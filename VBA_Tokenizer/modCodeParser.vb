Imports System.CodeDom
Imports System.Collections.Concurrent
Imports System.Text
Imports System.Text.RegularExpressions
Imports AVACONT_Core
Imports AVACONT_Core.modTypes
Imports AVACONT_Core.Logger

Public Module modCodeParser
    ' ==============================================================
    '                 BUILD SYMBOL DATABASE
    ' ==============================================================
    Public Sub ParseTextToObjects()
        ' === colecție surse: (Name, Type, Code, FCI)
        Dim sources As New List(Of (Name As String, Type As String, Code As String, FCI As Object))

        ' 1) Forms/Reports (cu CodeContent atașat)
        For Each kvp In GlobalModules
            Dim f = kvp.Value
            If Not String.IsNullOrWhiteSpace(f.CodeContent) Then
                sources.Add((f.Name, f.Type, f.CodeContent, f))
            End If
        Next
        For Each kv In GlobalFormsReports
            Dim f = kv.Value
            If Not String.IsNullOrWhiteSpace(f.CodeContent) Then
                sources.Add((f.Name, f.Type, f.CodeContent, f))
            End If
        Next

        Dim opts As New ParallelOptions With {.MaxDegreeOfParallelism = Environment.ProcessorCount}

        Dim runner As Action =
            Sub()
                If useParallel Then
                    Parallel.ForEach(sources, opts, Sub(src) ParseSingleCodeBlock(src.Name, src.Type, src.Code, src.FCI))
                Else
                    For Each src In sources
                        ParseSingleCodeBlock(src.Name, src.Type, src.Code, src.FCI)
                    Next
                End If
            End Sub

        runner()

    End Sub

    ' ==============================================================
    '                PARSARE pe o singură sursă
    ' ==============================================================
    Private Sub ParseSingleCodeBlock(currentModule As String, currentType As String, codeContent As String, formOrModule As Object)
        ' === Normalizează tip pentru decizii
        Dim isModule As Boolean = currentType.Equals("Module", StringComparison.OrdinalIgnoreCase)
        Dim isClassLike As Boolean =
            currentType.Equals("Class", StringComparison.OrdinalIgnoreCase) OrElse
            currentType.Equals("Forms", StringComparison.OrdinalIgnoreCase) OrElse
            currentType.Equals("Reports", StringComparison.OrdinalIgnoreCase)
        Dim condBlock As String = "" 'decide daca sunt in bloc #If...#Else...#End If
        Dim tmpCondBlock As String = ""
        Dim allLines = codeContent.Split({vbCrLf, vbLf}, StringSplitOptions.None)

        formOrModule.methods = New Dictionary(Of String, MethodInfo)
        formOrModule.Declares = New Dictionary(Of String, MethodInfo)
        formOrModule.Variables = New List(Of ParamInfo)
        formOrModule.Constants = New List(Of ParamInfo)

        'If currentModule = "clsDatabase" Then Stop

        'LogInfo($"Prelucrez metode / proprietati pentru {String.Join(":", currentModule, currentType)}")

        Try
            Dim merged = PreParseCode(codeContent)

            Dim currentClass As String = If(isClassLike, currentModule, "")
            Dim currentFunc As String = ""
            Dim beforeFirstFunc As Boolean = True
            Dim currentMethodInfo As MethodInfo = Nothing
            Dim allTokens As New List(Of Token)
            Dim isInType As Boolean
            Dim isInEnum As Boolean
            Dim isInMethod As Boolean = False
            Dim isInCondBlock As Boolean = False
            Dim methodCode As String = ""

            For Each tpl In merged
                Dim lineNumber = tpl.lineNum
                Dim L = tpl.text
                'If L.Contains("ControlHasProperty2") Then Stop

                If L.StartsWith("#If", StringComparison.OrdinalIgnoreCase) Then
                    isInCondBlock = True
                    Dim nextToken As String = L.Split(Space(1))(1)
                    Dim condToken As String
                    If nextToken.Equals("Not") Then
                        condToken = L.Split(Space(1))(2)
                    Else
                        condToken = nextToken
                    End If

                    condBlock = String.Join("#", condBlock, condToken)
                    'LogInfo(currentModule & " Bloc condițional început: " & condBlock)
                    Continue For

                ElseIf L.StartsWith("#Else", StringComparison.OrdinalIgnoreCase) Then
                    tmpCondBlock = condBlock.Split("#"c).Last()
                    Dim parts = condBlock.Split("#"c)
                    If parts.Length = 1 Then
                        condBlock = ""
                    Else
                        condBlock = String.Join("#", parts.Take(parts.Length - 1))
                    End If
                    tmpCondBlock = If(If(tmpCondBlock = "Win32", "Win64",
                                     If(tmpCondBlock = "Win64", "Win32",
                                     If(tmpCondBlock = "VBA6", "VBA7",
                                     If(tmpCondBlock = "VBA7", "VBA6",
                                     If(tmpCondBlock = "MAC", "",
                                        tmpCondBlock))))), tmpCondBlock)
                    condBlock = String.Join("#", condBlock, tmpCondBlock)
                    'LogInfo(currentModule & " Bloc condițional schimbat (Else): " & condBlock)
                    Continue For

                ElseIf L.StartsWith("#End") Then
                    If condBlock.Split("#"c).Length = 1 Then
                        condBlock = ""
                        isInCondBlock = False
                    Else
                        condBlock = condBlock.Substring(0, condBlock.LastIndexOf("#"c))
                    End If
                    'LogInfo(currentModule & " Bloc condițional încheiat: " & condBlock)
                    Continue For

                ElseIf {"Private", "Public"}.Any(Function(p) L.TrimStart().StartsWith(p & " Declare", StringComparison.OrdinalIgnoreCase)) Then
                    ParseDeclares(L, formOrModule, lineNumber, condBlock)
                    'LogInfo(currentModule & " Declare detectat: " & L)
                    Continue For

                ElseIf {"Global Const", "Public Const", "Private Const", "Const"}.Any(Function(p) L.TrimStart.StartsWith(p, StringComparison.OrdinalIgnoreCase)) And Not isInMethod Then
                    formOrModule.Constants.AddRange(ParseVariable(L, isClassLike, lineNumber, currentModule, RegexCache.RxConsts))
                    'LogInfo(currentModule & " Constantă detectată: " & L)
                    Continue For

                ElseIf {"Public Type", "Private Type", "Type"}.Any(Function(p) L.TrimStart.StartsWith(p, StringComparison.OrdinalIgnoreCase)) Then
                    isInType = True
                    'LogInfo(currentModule & " Început Type: " & L)
                    Continue For

                ElseIf {"Public Enum", "Private Enum", "Enum"}.Any(Function(p) L.TrimStart.StartsWith(p, StringComparison.OrdinalIgnoreCase)) Then
                    isInEnum = True
                    'LogInfo(currentModule & " Început Enum: " & L)
                    Continue For

                ElseIf {"Public Event", "Private Event", "Event"}.Any(Function(p) L.TrimStart.StartsWith(p, StringComparison.OrdinalIgnoreCase)) Then
                    ParseMethodDeclaration(L, formOrModule, currentMethodInfo, lineNumber, RegexCache.RxEvent, condBlock)
                    'LogInfo(currentModule & " Început Enum: " & L)
                    Continue For

                ElseIf RegexCache.RxMethodTester.IsMatch(L) Then
                    ParseMethodDeclaration(L, formOrModule, currentMethodInfo, lineNumber, IIf(L.Contains("Property "), RegexCache.RxProperty, RegexCache.RxHeader), condBlock)

                    'LogInfo(currentModule & " Metodă/Proprietate detectată: " & currentMethodInfo.Name)

                ElseIf L.Trim = "End Type" Then
                    isInType = False
                    'LogInfo(currentModule & " Sfârșit Type")
                    Continue For

                ElseIf L.Trim = "End Enum" Then
                    isInEnum = False
                    'LogInfo(currentModule & " Sfârșit Enum")
                    Continue For

                ElseIf isInEnum Or isInType Then
                    'TODO: 
                    ' - creeare doua noi clase: Types si Enums
                    ' - adaugare fiecare element din ele in clase de tipul paraminfo
                    Continue For

                ElseIf L.TrimStart.StartsWith("Attribute") Then
                    If isClassLike AndAlso currentMethodInfo IsNot Nothing Then
                        currentMethodInfo.isdefault = L.Contains("VB_UserMemId")
                        currentMethodInfo.isenumerable = L.Contains("VB_MemberFlags")
                    End If
                    Continue For

                ElseIf {"Dim", "Static", "Private", "Public"}.Any(Function(p) L.StartsWith(p, StringComparison.OrdinalIgnoreCase)) Then
                    If currentMethodInfo Is Nothing Then
                        formOrModule.variables.AddRange(ParseVariable(L, isClassLike, lineNumber, currentModule, RegexCache.RxGlobalVar))

                    Else
                        currentMethodInfo.variables.AddRange(ParseVariable(L, isClassLike, lineNumber, currentMethodInfo.name, RegexCache.RxLocalVar))
                    End If
                    'LogInfo(currentModule & " Variabilă detectată: " & L)
                    Continue For

                ElseIf {"Function", "Sub", "Property"}.Any(Function(p) L.TrimStart().StartsWith("End " & p, StringComparison.OrdinalIgnoreCase)) Then
                    If currentMethodInfo IsNot Nothing Then
                        currentMethodInfo.EndLine = lineNumber + 1

                        Dim startIdx = Math.Max(currentMethodInfo.StartLine - 1, 0)
                        Dim endIdx = Math.Min(currentMethodInfo.EndLine, allLines.Length)
                        Dim fullMethodName As String = currentMethodInfo.name & condBlock & IIf(currentMethodInfo.isproperty, currentMethodInfo.propertytype, "")

                        If endIdx >= startIdx Then
                            currentMethodInfo.CodeContent = String.Join(vbCrLf, allLines.Skip(startIdx).Take(endIdx - startIdx + 1))
                        Else
                            currentMethodInfo.CodeContent = ""
                        End If

                        ResolveVariableScopes(currentMethodInfo, formOrModule)
                        Try
                            formOrModule.Methods.Add(fullMethodName, currentMethodInfo)
                        Catch ex As Exception
                            Stop
                        End Try

                        currentMethodInfo = Nothing
                    Else
                        Stop
                        Throw New Exception("End without matching start at line " & lineNumber & " in module " & currentModule)
                    End If

                    currentFunc = ""
                    'LogInfo(currentModule & " Sfârșit metodă/proprietate la linia: " & lineNumber)
                    Continue For

                ElseIf currentMethodInfo IsNot Nothing Then
                    Dim tokenizedLine As New MethodLine With {
                        .LineNumber = lineNumber,
                        .Content = L,
                        .Tokens = GetTokensFromLine(L, lineNumber, currentFunc)
                    }

                    currentMethodInfo.MethodLines.Add(tokenizedLine)
                    'LogInfo(currentModule & " Linie metodă procesată: " & L)
                End If
            Next

            ' Dacă ultima metodă nu s-a închis, setează EndLine
            'If currentMethodInfo IsNot Nothing AndAlso merged.Count > 0 Then
            '    currentMethodInfo.EndLine = merged.Last().lineNum
            'End If

            'modExcelExporter.CollectModuleData(currentModule, currentType, formOrModule)
            LogInfo(ObjectCounter(formOrModule))
        Catch ex As Exception
            LogError($"ParseSingleCodeBlock({currentModule})", ex)

        End Try
    End Sub

    ' Helper: determină dacă un modul e Class-like (Form/Report/Class) sau Module standard
    'Public Function IsClassLikeModule(moduleName As String) As Boolean
    '    ' Verifică în GlobalFormsReports (Forms/Reports)
    '    If GlobalFormsReports.ContainsKey(moduleName) Then Return True

    '    ' Verifică în GlobalModules
    '    For Each kvp In GlobalModules
    '        Dim gm = TryCast(kvp.Value, ModuleContainer)
    '        If gm.Name.Equals(moduleName, StringComparison.OrdinalIgnoreCase) Then
    '            Return gm.Type.Equals("Class", StringComparison.OrdinalIgnoreCase) OrElse
    '                   gm.Type.Equals("Form", StringComparison.OrdinalIgnoreCase) OrElse
    '                   gm.Type.Equals("Report", StringComparison.OrdinalIgnoreCase)
    '        End If
    '    Next

    '    Return False

    'End Function

    ' ==============================================================
    '                HELPERS: VARIABLE DETECTION
    ' ==============================================================
    Private Function ObjectCounter(formOrModule As Object) As String
        Dim result As New StringBuilder()
        With formOrModule
            result.AppendLine($"Module: { .Name} ({ .Type})")
            result.AppendLine($"  Methods: { .Methods.Count}")
            result.AppendLine($"  Declares: { .Declares.Count}")
            result.AppendLine($"  Variables: { .Variables.Count}")
            result.AppendLine($"  Constants: { .Constants.Count}")
            Return result.ToString()
        End With
    End Function


    Private Sub ResolveVariableScopes(ByRef methodInfo As MethodInfo, formOrModule As Object)
        If methodInfo Is Nothing Then Exit Sub

        ' Construiesc dicționar cu TOATE variabilele accesibile
        Dim scopeVars As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        ' 1. Variabile la nivel de MODUL
        For Each v In formOrModule.Variables
            scopeVars(v.Name) = v.Type
        Next

        ' 2. Constante la nivel de MODUL
        For Each c In formOrModule.Constants
            scopeVars(c.Name) = c.Type
        Next

        ' 3. Parametrii metodei (suprascriu cele globale)
        For Each p In methodInfo.Parameters
            scopeVars(p.Name) = p.Type
        Next

        ' 4. Variabilele locale (suprascriu tot)
        For Each v In methodInfo.Variables
            scopeVars(v.Name) = v.Type
        Next

        ' 5. Parcurg toate liniile metodei și actualizez tokenii
        For Each ml In methodInfo.MethodLines
            For Each tok In ml.Tokens
                ' Variabilă directă
                If scopeVars.ContainsKey(tok.TokenString) Then
                    Select Case tok.TokenType
                        Case "identifier", "unknown"
                            tok.TokenType = "variable"
                            tok.DataType = scopeVars(tok.TokenString)
                        Case "object"
                            tok.DataType = scopeVars(tok.TokenString)
                    End Select
                End If

                ' Proprietate de obiect cunoscut
                If tok.Parent IsNot Nothing AndAlso scopeVars.ContainsKey(tok.Parent.TokenString) Then
                    If tok.TokenType = "property" OrElse tok.TokenType = "method" Then
                        tok.Parent.DataType = scopeVars(tok.Parent.TokenString)
                    End If
                End If
            Next
        Next
    End Sub

    Private Sub ParseDeclares(L As String, ByRef formOrModule As Object, lineNumber As Long, Optional condBlock As String = "")
        Dim paramsRaw As List(Of ParamInfo) = Nothing

        Dim h = RegexCache.RxDeclares.Match(L)
        Dim currentMethodInfo As Object = Nothing
        Dim access = If(h.Groups(1).Success, h.Groups(1).Value, "Public")
        Dim kind As String = ""
        Dim name As String = ""
        Dim returnType As String = ""
        Dim inner As String = ""
        Dim isPublic = False
        Dim fLib As String = ""
        Dim fAlias As String = ""

        Try
            If h.Groups(1).Success Then isPublic = h.Groups(1).Value = "Public"
            If h.Groups(3).Success Then kind = h.Groups(3).Value
            If h.Groups(4).Success Then name = h.Groups(4).Value
            If h.Groups(5).Success Then fLib = h.Groups(5).Value
            If h.Groups(6).Success Then fAlias = h.Groups(6).Value
            If h.Groups(7).Success Then inner = h.Groups(7).Value.Trim()
            If h.Groups(8).Success Then returnType = h.Groups(8).Value

            paramsRaw = New List(Of ParamInfo)
            Dim fullName = String.Join(".", formOrModule.Name, name) & condBlock

            Dim matches = RegexCache.RxParams.Matches(inner)
            For Each m As Match In matches
                If m.Success Then
                    ' Detectează tipul parametrilor din semnătura funcției
                    Dim pRef = m.Groups(1).Value
                    Dim isOptional = m.Groups(2).Value
                    Dim isParamArray = m.Groups(3).Value
                    Dim pName = m.Groups(4).Value
                    Dim pType = m.Groups(5).Value : If pType = "" Then pType = "Variant"
                    Dim pValue = m.Groups(6).Value

                    Dim prm As New ParamInfo With {
                        .Name = pName,
                        .Type = pType,
                        .IsByRef = pRef.Equals("ByRef", StringComparison.OrdinalIgnoreCase),
                        .IsOptional = isOptional.Equals("Optional", StringComparison.OrdinalIgnoreCase),
                        .IsParamArray = isParamArray.Equals("ParamArray", StringComparison.OrdinalIgnoreCase),
                        .IsGlobal = False
                    }
                    paramsRaw.Add(prm)
                End If
            Next

            currentMethodInfo = New MethodInfo With {
                .Name = name,
                .Signature = L,
                .ParentModule = formOrModule.name,
                .ParentType = formOrModule.Type,
                .StartLine = lineNumber,
                .EndLine = lineNumber,
                .MethodScope = access.ToLowerInvariant(),
                .MethodLines = New List(Of MethodLine),
                .IsFunction = kind.Equals("Function", StringComparison.OrdinalIgnoreCase),
                .ReturnType = returnType
            }

            For Each prm In paramsRaw
                prm.ParentMethod = currentMethodInfo
            Next

            currentMethodInfo.Parameters = paramsRaw

            formOrModule.Declares.add(fullName, currentMethodInfo)

        Catch ex As Exception
            LogError("ParseMethodDeclaration", ex)
            'Stop
        End Try
    End Sub

    Private Sub ParseMethodDeclaration(L As String, formOrModule As Object, ByRef currentMethodInfo As Object, lineNumber As Long, rxToUse As Regex, Optional condBlock As String = "")
        Dim paramsRaw As List(Of ParamInfo) = Nothing
        Dim access As String = "Public"
        Dim kind As String = ""
        Dim name As String = ""
        Dim returnType As String = ""
        Dim inner As String = ""
        Dim propertyType As String = ""
        'If L = "Public Property Get OpenForms() As Collection" Then Stop

        Try
            Dim m = rxToUse.Match(L)
            If Not m.Success Then Exit Sub

            ' Extrag grupurile în funcție de regex folosit
            If rxToUse Is RegexCache.RxEvent Then
                If m.Groups(1).Success Then access = m.Groups(1).Value
                kind = "Event"
                name = m.Groups(2).Value
                inner = m.Groups(3).Value.Trim()
                returnType = ""

            ElseIf rxToUse Is RegexCache.RxProperty Then
                access = If(m.Groups(1).Success, m.Groups(1).Value, "Public")
                kind = "Property"
                propertyType = m.Groups(2).Value
                name = m.Groups(3).Value
                inner = m.Groups(4).Value.Trim()
                returnType = If(m.Groups(5).Success, m.Groups(5).Value, "")

            ElseIf rxToUse Is RegexCache.RxHeader Then
                access = If(m.Groups(1).Success, m.Groups(1).Value, "Public")
                kind = m.Groups(2).Value  ' Function sau Sub
                name = m.Groups(3).Value
                inner = m.Groups(4).Value.Trim()
                returnType = If(m.Groups(5).Success, m.Groups(5).Value, "")
            End If

            paramsRaw = New List(Of ParamInfo)
            Dim fullName = String.Join(".", formOrModule.Name, name) & If(condBlock <> "", "#" & condBlock, "")

            Dim matches = RegexCache.RxParams.Matches(inner)
            Dim tokenizedLine As New MethodLine With {.LineNumber = lineNumber, .Content = L}

            For Each pm As Match In matches
                If pm.Success Then
                    Dim pRef = pm.Groups(1).Value
                    Dim isOptional = pm.Groups(2).Value
                    Dim isParamArray = pm.Groups(3).Value
                    Dim pName = pm.Groups(4).Value
                    Dim pType = pm.Groups(5).Value : If pType = "" Then pType = "Variant"
                    Dim pValue = pm.Groups(6).Value

                    Dim prm As New ParamInfo With {
                        .Name = pName,
                        .Type = pType,
                        .IsByRef = pRef.Equals("ByRef", StringComparison.OrdinalIgnoreCase),
                        .IsOptional = isOptional.Equals("Optional", StringComparison.OrdinalIgnoreCase),
                        .IsParamArray = isParamArray.Equals("ParamArray", StringComparison.OrdinalIgnoreCase),
                        .IsGlobal = False
                    }

                    paramsRaw.Add(prm)

                    Dim paramToken As New Token With {
                        .Context = name,
                        .IsLeftSide = False,  ' parametrii nu sunt LHS
                        .LineNumber = lineNumber,
                        .NextSymbol = "As",
                        .TokenString = prm.Name,
                        .TokenType = "parameter"  ' nu "Variable"
                    }

                    tokenizedLine.Tokens.Add(paramToken)

                    If Not String.IsNullOrEmpty(prm.Type) Then
                        Dim typeToken As New Token With {
                            .Context = name,
                            .IsLeftSide = False,
                            .LineNumber = lineNumber,
                            .NextSymbol = "",
                            .TokenString = prm.Type,
                            .TokenType = "type",  ' lowercase pentru consistență
                            .Parent = paramToken  ' legătura cu parametrul
                        }

                        tokenizedLine.Tokens.Add(typeToken)

                        ' Salvez tipul în parametru pentru referință
                        paramToken.DataType = prm.Type
                    End If
                End If
            Next

            currentMethodInfo = New MethodInfo With {
                .Name = name,
                .Signature = L,
                .ParentModule = formOrModule.name,
                .ParentType = formOrModule.Type,
                .StartLine = lineNumber,
                .EndLine = If(kind = "Event", lineNumber, 0),
                .MethodScope = access.ToLowerInvariant(),
                .MethodLines = New List(Of MethodLine),
                .IsFunction = kind.Equals("Function", StringComparison.OrdinalIgnoreCase),
                .IsProperty = kind.Equals("Property", StringComparison.OrdinalIgnoreCase),
                .ReturnType = returnType,
                .PropertyType = propertyType
            }

            For Each prm In paramsRaw
                prm.ParentMethod = currentMethodInfo
            Next

            currentMethodInfo.Parameters = paramsRaw
            If tokenizedLine.Tokens.Count > 0 Then currentMethodInfo.TokenizedLine = tokenizedLine

        Catch ex As Exception
            LogError("ParseMethodDeclaration", ex)
        End Try
    End Sub

    Private Function ParseVariable(L As String, isClassLike As Boolean, lineNumber As Integer, context As String, rxToUse As Regex) As List(Of ParamInfo)
        Dim matches = rxToUse.Matches(L)
        Dim isGlobal As Boolean = False
        Dim result As New List(Of ParamInfo)

        Try
            For Each m As Match In matches
                If m.Success Then
                    ' Extrag valorile în funcție de regex
                    Dim access As String = ""
                    Dim isWithEvents As Boolean = False
                    Dim pName As String = ""
                    Dim pType As String = ""
                    Dim pValue As String = ""

                    ' Determin ce regex e folosit și extrag grupurile corespunzător
                    If rxToUse Is RegexCache.RxGlobalVar Then
                        access = m.Groups(1).Value
                        isWithEvents = m.Groups(2).Value <> ""
                        pName = m.Groups(3).Value
                        pType = m.Groups(4).Value
                        If access = "" Then
                            isGlobal = Not isClassLike
                        Else
                            isGlobal = access.Contains("Public") And Not isClassLike
                        End If

                    ElseIf rxToUse Is RegexCache.RxLocalVar Then
                        pName = m.Groups(1).Value
                        pType = m.Groups(2).Value

                    ElseIf rxToUse Is RegexCache.RxConsts Then
                        access = m.Groups(1).Value
                        pName = m.Groups(3).Value
                        pType = m.Groups(4).Value
                        pValue = m.Groups(5).Value
                        If access = "" Then
                            isGlobal = Not isClassLike
                        Else
                            isGlobal = access.Contains("Public") And Not isClassLike
                        End If
                    End If

                    ' Creez tokeni
                    Dim tokens As New List(Of Token)

                    ' Token pentru numele variabilei
                    Dim nameToken As New Token With {
                        .LineNumber = lineNumber,
                        .TokenString = pName,
                        .TokenType = "variable",  ' consistent
                        .Context = context,
                        .Parent = Nothing,
                        .NextSymbol = If(String.IsNullOrEmpty(pType), "", "As"),
                        .IsLeftSide = True  ' variabilele în declarații sunt LHS
                    }
                    tokens.Add(nameToken)

                    ' Tokenizez tipul
                    If Not String.IsNullOrEmpty(pType) Then
                        Dim typeParts = pType.Split("."c)
                        Dim prevToken As Token = nameToken

                        For i = 0 To typeParts.Length - 1
                            Dim part = typeParts(i)
                            Dim nextSym As String = If(i < typeParts.Length - 1, ".", "")

                            Dim typeToken As New Token With {
                                .LineNumber = lineNumber,
                                .TokenString = part,
                                .TokenType = If(i = 0, "type", "property"), ' primul e tip, restul proprietăți
                                .Context = context,
                                .Parent = prevToken,
                                .NextSymbol = nextSym,
                                .IsLeftSide = False
                            }
                            tokens.Add(typeToken)
                            prevToken = typeToken
                        Next

                        ' Salvez tipul în token-ul variabilei
                        nameToken.DataType = pType
                    End If

                    Dim prm As New ParamInfo With {
                        .Name = pName,
                        .IsGlobal = isGlobal,
                        .IsPublic = rxToUse IsNot RegexCache.RxParams AndAlso (access = "Public" Or access = "" And Not isClassLike),
                        .Type = pType,
                        .Value = pValue,
                        .Tokens = tokens
                    }

                    result.Add(prm)
                End If
            Next

            Return result

        Catch ex As Exception
            LogError("ParseVariable", ex)
            Return New List(Of ParamInfo)
        End Try
    End Function

    Private Function PreParseCode(codeContent As String) As List(Of (lineNum As Integer, text As String))
        ' === Preproc: unește liniile cu underscore și numerotează ===
        Dim rawLines = codeContent.Split({vbCrLf, vbLf}, StringSplitOptions.None).ToList()
        Dim merged As New List(Of (lineNum As Integer, text As String))
        Dim buffer As String = ""
        Dim ln As Integer = 1, startNum As Integer = 1
        Dim isContinued = False

        Try
            For Each raw In rawLines
                Dim L = raw.TrimEnd()
                'If L.Contains("Function InitGDIP()") Then Stop
                'If L.Contains("BCryptHashData") Then Stop

                'If ln = 1530 Then Stop

                ' === ignora comentarii, linii goale, Option ===
                If L.TrimStart().StartsWith("'@@") Then
                    Continue For
                ElseIf L = "" OrElse L.TrimStart().StartsWith("'") Then
                    ln += 1 : Continue For
                ElseIf L.TrimStart().StartsWith("Option ") Then
                    ln += 1 : Continue For
                End If

                ' === concatenare pentru linii cu underscore ===
                If L.EndsWith("_") Then
                    If Not isContinued Then startNum = ln
                    buffer &= " " & L.TrimEnd("_"c, "&"c).Trim()
                    isContinued = True
                ElseIf isContinued Then
                    buffer &= " " & L.TrimEnd("_"c, "&"c).Trim()
                    ' finalizează linia continuată
                    For Each part In RegexCache.RxSplitColon.Split(buffer)
                        Dim cleaned = CleanLineForShits(part.Trim())
                        If cleaned <> "" Then merged.Add((startNum, cleaned))
                    Next
                    buffer = ""
                    isContinued = False
                    startNum = 0
                Else
                    ' === linie normală ===
                    'imparte linia dupa : pentru a-i putea procesa tokenii
                    For Each part In RegexCache.RxSplitColon.Split(L)
                        'daca ramane un singur cuvant in element, nu poate fi altceva deca label destinatie pentru goto
                        If part.Split(" "c).Length = 1 Then
                            'ii pun inapoi:
                            merged.Add((ln, part & ":"))
                        Else
                            Dim cleaned = CleanLineForShits(part.Trim())
                            If cleaned <> "" Then merged.Add((ln, cleaned))
                        End If
                    Next
                End If

                ln += 1
            Next

            ' === rest dacă a rămas ceva în buffer ===
            If buffer <> "" Then
                For Each part In RegexCache.RxSplitColon.Split(buffer)
                    Dim cleaned = CleanLineForShits(part.Trim())
                    If cleaned <> "" Then merged.Add((startNum, cleaned))
                Next
            End If

            Return merged

        Catch ex As Exception
            LogError("PreParseCode", ex)
            Return New List(Of (lineNum As Integer, text As String))
        End Try
    End Function

    Private Function ParseClassField(line As String, currentModule As String, currentClass As String) As VariableInfo
        Dim m = RegexCache.RxClassField.Match(line)
        If Not m.Success Then Return Nothing

        Return New VariableInfo With {
            .Name = m.Groups("n").Value,
            .TypeName = If(m.Groups("type").Success, m.Groups("type").Value, ""),
            .CallerModule = currentModule,
            .ScopeFunc = currentClass,
            .IsWithEvents = m.Groups("we").Success
        }
    End Function

    Private Function CleanLineForShits(line As String) As String
        ' Elimina mai mult de un spatiu consecutiv
        Dim noSpaces = RegexCache.RxSpaces.Replace(line, " ")
        ' Elimină stringurile și comentariile pentru a facilita parsarea
        Dim noStrings = If(RegexCache.RxComments.IsMatch(noSpaces),
                   RegexCache.RxComments.Match(noSpaces).Groups(1).Value,
                   noSpaces)
        ' Elimina situatia " & " din stringuri
        Dim noAmpersands = noStrings.Replace(""" & """, "")

        Return noAmpersands.Trim()
    End Function
End Module