'PROJECT NAME: VBA_TOKENIZER
'FILE DESCRIPTION: Module pentru parsarea codului VBA în obiecte (metode, variabile, constante, declarații)
'PATH: VBA_TOKENIZER/modCodeParser.vb
Imports System.Collections.Concurrent
Imports System.Text.RegularExpressions
Imports System.Threading
Imports VBA_CORE
Imports VBA_CORE.Logger
Imports VBA_CORE.modTypes

Public Module modCodeParser
    Private pCurrentModule As New ThreadLocal(Of String)(Function() "")
    Private pCurrentType As New ThreadLocal(Of String)(Function() "")
    Private pCurrentMethod As New ThreadLocal(Of String)(Function() "")
    Private pCurrentEvent As New ThreadLocal(Of String)(Function() "")
    ' ==============================================================
    '                 BUILD SYMBOL DATABASE
    ' ==============================================================
    Public Sub ParseTextToObjects()
        ' === colecție surse: (Name, Type, Code, FCI)
        GlobalSources = New ConcurrentBag(Of (Name As String, Type As String, Code As String, FCI As ModuleContainer))

        Try
            ' 1) Forms/Reports (cu CodeContent atașat)
            For Each kvp In GlobalModules
                Dim f = kvp.Value
                If Not String.IsNullOrWhiteSpace(f.CodeContent) Then
                    GlobalSources.Add((f.Name, f.Type, f.CodeContent, f))
                End If
            Next
            For Each kv In GlobalFormsReports
                Dim f = kv.Value
                If Not String.IsNullOrWhiteSpace(f.CodeContent) Then
                    GlobalSources.Add((f.Name, f.Type, f.CodeContent, f))
                End If
            Next

            Dim opts As New ParallelOptions With {.MaxDegreeOfParallelism = Environment.ProcessorCount}

            Dim runner As Action =
                Sub()
                    If modGlobals.Global_UseParallel Then
                        Parallel.ForEach(GlobalSources, opts, Sub(src) ParseSingleCodeBlock(src.Name, src.Type, src.Code, src.FCI))
                    Else
                        For Each src In GlobalSources
                            ParseSingleCodeBlock(src.Name, src.Type, src.Code, src.FCI)
                        Next
                    End If
                End Sub

            runner()

        Catch ex As Exception
            LogError(Global_ExportDir & " ParseTextToObjects", ex)
        End Try
    End Sub

    ' ==============================================================
    '                PARSARE pe o singură sursă
    ' ==============================================================
    Private Sub ParseSingleCodeBlock(currentModule As String, currentType As String, codeContent As String, MCI As ModuleContainer)
        ' === Salvez context curent pentru logging
        pCurrentModule.Value = currentModule
        pCurrentType.Value = currentType
        pCurrentMethod.Value = ""
        'If currentModule = "clsDatabase" Then Stop

        ' === Normalizează tip pentru decizii
        Dim isModule As Boolean = currentType.Equals("Module", StringComparison.OrdinalIgnoreCase)
        Dim isClassLike As Boolean =
            currentType.Equals("Class", StringComparison.OrdinalIgnoreCase) OrElse
            currentType.Equals("Forms", StringComparison.OrdinalIgnoreCase) OrElse
            currentType.Equals("Reports", StringComparison.OrdinalIgnoreCase)
        Dim condBlock As String = "" 'decide daca sunt in bloc #If...#Else...#End If
        Dim tmpCondBlock As String = ""
        Dim allLines = codeContent.Split({vbCrLf, vbLf}, StringSplitOptions.None)
        Dim withStack As New Stack(Of MethodLine)
        Dim enumOrType As EnumOrType = Nothing

        MCI.Methods = New Dictionary(Of String, MethodInfo)
        MCI.Declares = New Dictionary(Of String, MethodInfo)
        MCI.Variables = New List(Of ParamInfo)
        MCI.Constants = New List(Of ParamInfo)
        MCI.Events = New Dictionary(Of String, MethodInfo)
        MCI.Types = New List(Of EnumOrType)
        MCI.Enums = New List(Of EnumOrType)

        'If currentModule = "clsDatabase" Then Stop

        Try
            Dim merged = PreParseCode(codeContent, currentModule)

            Dim currentClass As String = If(isClassLike, currentModule, "")
            Dim currentFunc As String = ""
            Dim beforeFirstFunc As Boolean = True
            Dim tokenizedLine As MethodLine
            Dim currentMethodInfo As MethodInfo = Nothing
            Dim allTokens As New List(Of Token)
            Dim isInType As Boolean
            Dim isInEnum As Boolean
            Dim isInMethod As Boolean = False
            Dim isInCondBlock As Boolean = False
            Dim methodCode As String = ""
            Dim localLineNum As Long = 0
            Dim methodLineNum As Long = 0

            For Each tpl In merged
                Dim lineNumber = tpl.lineNum
                Dim L = tpl.text
                Dim rawL = tpl.rawLine

                localLineNum += 1

                tokenizedLine = New MethodLine With {
                    .LocalLineNumber = localLineNum,
                    .LineNumber = lineNumber,
                    .Content = L,
                    .OriginalContent = rawL,
                    .IsInEnumBlock = isInEnum,
                    .IsInTypeBlock = isInType
                }

                If withStack.Count > 0 Then
                    tokenizedLine.IsInWithBlock = True
                    tokenizedLine.WithBlockContext = withStack.Peek()
                    tokenizedLine.WithBlockDepth = withStack.Count + 1
                End If

                If L.StartsWith("#If", StringComparison.OrdinalIgnoreCase) Then
                    'inceput bloc condițional
                    isInCondBlock = True
                    Dim nextToken As String = L.Split(Space(1))(1)
                    Dim condToken As String
                    If nextToken.Equals("Not") Then
                        condToken = L.Split(Space(1))(2)
                    Else
                        condToken = nextToken
                    End If

                    condBlock = String.Join("#", condBlock, condToken)
                    LogInfoLocal(" Conditional block (Start): " & condBlock, 2)
                    'continue for

                ElseIf L.StartsWith("#Else", StringComparison.OrdinalIgnoreCase) Then
                    'schimbare condițional
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
                    LogInfoLocal(" Conditional block (Else): " & condBlock, 2)
                    'continue for

                ElseIf L.StartsWith("#End") Then
                    'sfârșit bloc condițional
                    If condBlock.Split("#"c).Length = 1 Then
                        condBlock = ""
                        isInCondBlock = False
                    Else
                        condBlock = condBlock.Substring(0, condBlock.LastIndexOf("#"c))
                    End If
                    LogInfoLocal(" Conditional block (End): " & condBlock, 2)
                    'continue for

                ElseIf {"Private", "Public"}.Any(Function(p) L.TrimStart().StartsWith(p & " Declare", StringComparison.OrdinalIgnoreCase)) Then
                    ' Declarare functii API
                    tokenizedLine.IsDeclarationLine = True
                    ParseDeclares(L, MCI, lineNumber, condBlock)
                    MCI.Lines.Add(tokenizedLine)
                    LogInfoLocal(" API: " & L, 2)
                    Continue For

                ElseIf {"Global Const", "Public Const", "Private Const", "Const"}.Any(Function(p) L.TrimStart.StartsWith(p, StringComparison.OrdinalIgnoreCase)) And Not isInMethod Then
                    ' Constante la nivel de modul
                    MCI.Constants.AddRange(ParseVariable(L, isClassLike, lineNumber, currentModule, RegexCache.RxConsts, currentModule, rawL, localLineNum, 0))
                    tokenizedLine.IsDeclarationLine = True
                    LogInfoLocal(" Constant: " & L, 2)
                    'continue for

                ElseIf {"Public Event", "Private Event", "Event"}.Any(Function(p) L.TrimStart.StartsWith(p, StringComparison.OrdinalIgnoreCase)) AndAlso currentMethodInfo Is Nothing Then
                    ' Declarare eveniment
                    tokenizedLine.IsEventLine = True
                    ParseMethodDeclaration(L, MCI, currentMethodInfo, lineNumber, RegexCache.RxEvent, rawL, condBlock, localLineNum, 0)
                    LogInfoLocal(" Event: " & L, 2)
                    MCI.Events.Add(currentMethodInfo.Name, currentMethodInfo)
                    MCI.Lines.Add(tokenizedLine)
                    currentMethodInfo = Nothing
                    Continue For

                ElseIf RegexCache.RxMethodTester.IsMatch(L) Then
                    'inceput metodă/proprietate
                    methodLineNum = 0
                    ParseMethodDeclaration(L, MCI, currentMethodInfo, lineNumber, IIf(L.Contains("Property "), RegexCache.RxProperty, RegexCache.RxHeader), rawL, condBlock, localLineNum, methodLineNum)
                    pCurrentMethod.Value = currentMethodInfo.Name
                    tokenizedLine.StartsMethodBlock = True
                    LogInfoLocal(" Method/Property:" & L, 2)

                ElseIf {"Public Type", "Private Type", "Type"}.Any(Function(p) L.TrimStart.StartsWith(p, StringComparison.OrdinalIgnoreCase)) Then
                    ' Început Type
                    isInType = True
                    tokenizedLine.StartsTypeBlock = True
                    enumOrType = New EnumOrType With {
                        .Name = RegexCache.RxEnumTypeStart.Match(L).Groups(1).Value,
                        .Type = "Type",
                        .IsPublic = Not L.TrimStart().StartsWith("Private", StringComparison.OrdinalIgnoreCase),
                        .IsGlobal = MCI.Type.Equals("Module", StringComparison.OrdinalIgnoreCase)
                    }

                    LogInfoLocal(" Type: " & L, 2)

                ElseIf L.Trim = "End Type" Then
                    ' Sfârșit Type
                    isInType = False
                    tokenizedLine.EndsTypeBlock = True
                    tokenizedLine.IsInTypeBlock = False

                    MCI.Types.Add(enumOrType)

                    LogInfoLocal($"Type '{enumOrType.Name}' parsed with {enumOrType.Members.Count} members.", 2)

                    enumOrType = Nothing
                    'continue for

                ElseIf {"Public Enum", "Private Enum", "Enum"}.Any(Function(p) L.TrimStart.StartsWith(p, StringComparison.OrdinalIgnoreCase)) Then
                    ' Început Enum
                    isInEnum = True
                    tokenizedLine.StartsEnumBlock = True
                    enumOrType = New EnumOrType With {
                        .Name = RegexCache.RxEnumTypeStart.Match(L).Groups(1).Value,
                        .Type = "Enum",
                        .IsPublic = Not L.TrimStart().StartsWith("Private", StringComparison.OrdinalIgnoreCase),
                        .IsGlobal = MCI.Type.Equals("Module", StringComparison.OrdinalIgnoreCase)
                    }

                    LogInfoLocal(" Enum: " & L, 2)

                ElseIf L.Trim = "End Enum" Then
                    ' Sfârșit Enum
                    tokenizedLine.EndsEnumBlock = True
                    tokenizedLine.IsInEnumBlock = False
                    isInEnum = False

                    MCI.Enums.Add(enumOrType)

                    LogInfoLocal($"Enum '{enumOrType.Name}' parsed with {enumOrType.Members.Count} members.", 2)

                    enumOrType = Nothing
                    'continue for

                ElseIf isInEnum Or isInType Then
                    ' Prelucrare in interior Type/Enum
                    tokenizedLine.IsInEnumBlock = isInEnum
                    tokenizedLine.IsInTypeBlock = isInType
                    ParseEnumTypeLine(L, enumOrType, IIf(isInEnum, RegexCache.RxEnumMember, RegexCache.RxTypeMember))

                ElseIf L.TrimStart.StartsWith("Attribute") Then
                    ' Atribute VBA speciale (ignorate la tokenizare, dar utile pentru meta-informații)
                    If isClassLike AndAlso currentMethodInfo IsNot Nothing Then
                        currentMethodInfo.IsDefault = L.Contains("VB_UserMemId")
                        currentMethodInfo.IsEnumerable = L.Contains("VB_MemberFlags")
                    End If
                    LogInfoLocal(" Attribute:" & L, 2)

                    'continue for

                ElseIf {"Dim", "Static", "Private", "Public"}.Any(Function(p) L.StartsWith(p, StringComparison.OrdinalIgnoreCase)) Then
                    ' Variabile locale sau globale
                    If currentMethodInfo Is Nothing Then
                        ' Variabile la nivel de metoda
                        MCI.Variables.AddRange(ParseVariable(L, isClassLike, lineNumber, currentModule, RegexCache.RxGlobalVar, currentModule, rawL, localLineNum, methodLineNum))
                    Else
                        ' Variabile la nivel de modul
                        currentMethodInfo.Variables.AddRange(ParseVariable(L, isClassLike, lineNumber, currentMethodInfo.Name, RegexCache.RxLocalVar, currentModule, rawL, localLineNum, localLineNum))
                    End If
                    tokenizedLine.IsDeclarationLine = True
                    LogInfoLocal(" Variable: " & L, 2)
                    'continue for

                ElseIf {"Function", "Sub", "Property"}.Any(Function(p) L.TrimStart().StartsWith("End " & p, StringComparison.OrdinalIgnoreCase)) Then
                    'sfârșit metodă/proprietate
                    If currentMethodInfo IsNot Nothing Then
                        currentMethodInfo.EndLine = lineNumber + 1
                        tokenizedLine.EndsMethodBlock = True

                        Dim startIdx = Math.Max(currentMethodInfo.StartLine - 1, 0)
                        Dim endIdx = Math.Min(currentMethodInfo.EndLine, allLines.Length)
                        Dim fullMethodName As String = currentMethodInfo.Name & condBlock

                        If TypeOf currentMethodInfo Is PropertyInfo Then
                            Dim p = DirectCast(currentMethodInfo, PropertyInfo)
                            fullMethodName &= p.PropertyType
                        End If

                        If endIdx >= startIdx Then
                            currentMethodInfo.CodeContent = String.Join(vbCrLf, allLines.Skip(startIdx).Take(endIdx - startIdx + 1).SkipWhile(Function(ln) String.IsNullOrWhiteSpace(ln)))
                        Else
                            currentMethodInfo.CodeContent = ""
                        End If

                        'tokenizedLine = New MethodLine With {.LineNumber = lineNumber, .Content = L}
                        currentMethodInfo.MethodLines.Add(tokenizedLine)
                        MCI.Lines.Add(tokenizedLine)

                        'ResolveVariableScopes(currentMethodInfo, MCI)

                        LogInfoLocal(" Line parsed: " & L, 2)

                        ' === Finalizare metodă ===
                        Try
                            MCI.Methods.Add(fullMethodName, currentMethodInfo)

                            ' === Clasificare pe tip ===
                            If TypeOf currentMethodInfo Is FunctionInfo Then
                                MCI.Functions(fullMethodName) = DirectCast(currentMethodInfo, FunctionInfo)

                            ElseIf TypeOf currentMethodInfo Is PropertyInfo Then
                                MCI.Properties(fullMethodName) = DirectCast(currentMethodInfo, PropertyInfo)

                            ElseIf currentMethodInfo.Name.StartsWith("Event", StringComparison.OrdinalIgnoreCase) Then
                                MCI.Events(fullMethodName) = currentMethodInfo
                            End If

                        Catch ex As Exception
                            LogError($"Duplicate or invalid method in {currentModule}: {currentMethodInfo.Name}", ex)
                        End Try

                        currentMethodInfo = Nothing
                        methodLineNum = 0
                        Continue For

                    Else
                        Stop
                        Throw New Exception("End without matching start at line " & lineNumber & " in module " & currentModule)
                    End If

                    currentFunc = ""
                    LogInfoLocal(" End method: " & lineNumber, 2)
                    'continue for

                ElseIf L.TrimStart().StartsWith("With ", StringComparison.OrdinalIgnoreCase) Then
                    ' ===== Începutul unui bloc With =====
                    Dim withExpr As String = L.Substring(L.IndexOf("With", StringComparison.OrdinalIgnoreCase) + 4).Trim()

                    tokenizedLine.StartsWithBlock = True

                    withStack.Push(tokenizedLine)
                    LogInfoLocal($" With block START: '{withExpr}' (depth={withStack.Count})", 2)
                    'Continue For

                ElseIf L.TrimStart().StartsWith("End With", StringComparison.OrdinalIgnoreCase) Then
                    ' ===== Sfârșitul unui bloc With =====
                    tokenizedLine.EndsWithBlock = True
                    If withStack.Count > 0 Then
                        Dim closedWith = withStack.Pop()
                        LogInfoLocal($" With block END: '{closedWith}' (depth={withStack.Count})", 2)
                    Else
                        LogInfoLocal(" Warning: End With without matching With", 2)
                    End If
                    'Continue For
                Else

                End If

                ' === Procesare linie normală (în interior sau în afara unei metode)
                If currentMethodInfo IsNot Nothing Then
                    methodLineNum += 1
                    currentMethodInfo.MethodLines.Add(tokenizedLine)
                End If

                MCI.Lines.Add(tokenizedLine)


                LogInfoLocal(" Line parsed: " & L, 2)
            Next

            'modExcelExporter.CollectModuleData(currentModule, currentType, MCI)
            LogInfoLocal(ObjectCounter(MCI), 1)

            BuildModuleSymbols(MCI)
            AddToGlobalSymbols(MCI)

        Catch ex As Exception
            LogError($"ParseSingleCodeBlock({currentModule})", ex)

        End Try
    End Sub

    ' ==============================================================
    '                HELPERS
    ' ==============================================================
    Private Function ObjectCounter(formOrModule As Object) As String
        Dim totalTokens As Long = 0
        For Each m In formOrModule.Methods.Values
            For Each ml In m.MethodLines
                totalTokens += ml.Tokens.Count
            Next
        Next

        Return $"Methods: {formOrModule.Methods.Count}; " &
           $"Declares: {formOrModule.Declares.Count}; " &
           $"Variables: {formOrModule.Variables.Count}; " &
           $"Constants: {formOrModule.Constants.Count}; " &
           $"Tokens: {totalTokens}"
    End Function

    Private Sub LogInfoLocal(ByVal msg As String, level As Integer, Optional force As Boolean = False)
        If pCurrentMethod.Value <> "" Then
            msg = $"({pCurrentType.Value}) [{pCurrentModule.Value}].[{pCurrentMethod.Value}] : {msg}"
        Else
            msg = $"({pCurrentType.Value}) [{pCurrentModule.Value}] : {msg}"
        End If

        LogInfo(msg, level, force)
    End Sub

    Private Sub ParseEnumTypeLine(L As String, ByRef enumOrType As EnumOrType, rxToUse As Regex)
        Try
            Dim raw = L.Trim()

            ' === 1️⃣ extrage context din linia parinte ===
            Dim isEnum As Boolean = rxToUse Is RegexCache.RxEnumMember

            ' ex: Public Enum CustomerStatus → numele = CustomerStatus
            Dim parentName As String = enumOrType.Name

            ' === 2️⃣ identifică nume + tip / valoare ===
            Dim fieldName As String = ""
            Dim fieldType As String = ""
            Dim fieldValue As String = ""
            Dim m = rxToUse.Match(raw)

            If rxToUse Is RegexCache.RxEnumMember Then
                If m.Success Then
                    fieldName = m.Groups("name").Value
                    fieldValue = If(m.Groups("val").Success, m.Groups("val").Value, "")
                    fieldType = "Long"
                End If

            Else
                If m.Success Then
                    fieldName = m.Groups("name").Value
                    fieldType = m.Groups("type").Value
                End If
            End If

            If String.IsNullOrEmpty(fieldName) Then Exit Sub

            ' === 3️⃣ compune calificativul complet
            Dim qualifiedName As String = $"{parentName}.{fieldName}"

            ' === 4️⃣ creează simbolul și îl adaugă în container ===
            ' === Creează simbolul pentru membru Enum/Type ===
            Dim memberTokenType As TokenTypeEnum = If(isEnum, TokenTypeEnum.enum_decl, TokenTypeEnum.type_field_decl)
            Dim memberToken As New Token With {.TokenString = fieldName, .TokenType = memberTokenType, .LineNumber = 0}

            Dim member As New ParamInfo With {
                .Name = fieldName,
                .Type = fieldType,
                .IsGlobal = enumOrType.IsGlobal,
                .IsPublic = enumOrType.IsPublic,
                .QualifiedName = enumOrType.Name & "." & fieldName,
                .Tokens = New List(Of Token) From {memberToken}
            }

            If isEnum Then
                member.Value = fieldValue
            End If

            enumOrType.Members.Add(member)

            LogInfoLocal($"Parsed {If(isEnum, "Enum", "Type")} member: {qualifiedName} ({fieldType})", 2)

        Catch ex As Exception
            LogError($"ParseEnumTypeLine({enumOrType.Name})", ex)
        End Try
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

            If kind.Equals("Function", StringComparison.OrdinalIgnoreCase) Then
                currentMethodInfo = New FunctionInfo With {
                    .IsFunction = True,
                    .ReturnType = returnType,
                    .ReturnObject = Nothing
                }
            Else
                currentMethodInfo = New MethodInfo()
            End If

            With currentMethodInfo
                .Name = name
                .Signature = L
                .ParentModule = formOrModule.Name
                .ParentType = formOrModule.Type
                .StartLine = lineNumber
                .EndLine = lineNumber
                .MethodScope = access.ToLowerInvariant()
                .MethodLines = New List(Of MethodLine)
                .Parameters = paramsRaw
            End With

            For Each prm In paramsRaw
                prm.ParentMethod = currentMethodInfo
            Next

            currentMethodInfo.Parameters = paramsRaw

            formOrModule.Declares.add(fullName, currentMethodInfo)

            LogInfoLocal($" Parsed declare: {currentMethodInfo.Name} with {paramsRaw.Count} params.", 2)
        Catch ex As Exception
            LogError($" -> ({formOrModule.name}) ParseDeclare", ex)
            'Stop
        End Try
    End Sub

    Private Sub ParseMethodDeclaration(L As String, formOrModule As Object, ByRef currentMethodInfo As Object, lineNumber As Long, rxToUse As Regex, rawLine As String, condBlock As String, localLineNumber As Integer, methodLineNumber As Integer)
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
            Dim tokenizedLine As New MethodLine With {.LineNumber = lineNumber, .Content = L, .LocalLineNumber = localLineNumber}

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

                    ' 🔹 Adaug token explicit (util pentru analiză ulterioară)
                    Dim paramToken As New Token With {
                        .TokenString = pName,
                        .TokenType = TokenTypeEnum.parameter_decl,   ' 🔹 clarificare semnătură
                        .LineNumber = lineNumber,
                        .LocalLineNumber = localLineNumber,
                        .MethodLineNumber = methodLineNumber
                    }

                    prm.Tokens.Add(paramToken)
                    paramsRaw.Add(prm)

                End If
            Next

            If kind.Equals("Function", StringComparison.OrdinalIgnoreCase) Then
                currentMethodInfo = New FunctionInfo With {
                    .IsFunction = True,
                    .ReturnType = returnType,
                    .ReturnObject = Nothing
                }

            ElseIf kind.Equals("Property", StringComparison.OrdinalIgnoreCase) Then
                currentMethodInfo = New PropertyInfo With {
                    .IsProperty = True,
                    .PropertyType = propertyType,
                    .ReturnType = returnType,
                    .ReturnObject = Nothing
                }

            Else
                currentMethodInfo = New MethodInfo()
            End If

            ' === comun pentru toate
            With currentMethodInfo
                .Name = name
                .Signature = L
                .ParentModule = formOrModule.Name
                .ParentType = formOrModule.Type
                .StartLine = lineNumber
                .EndLine = If(kind = "Event", lineNumber, 0)
                .MethodScope = access.ToLowerInvariant()
                .MethodLines = New List(Of MethodLine)
                .Parameters = paramsRaw
            End With

            ' === 🔹 Detectăm evenimente implicite VBA (Form_Click, Text1_AfterUpdate etc.)
            Try
                Dim handlerObj As String = Nothing
                Dim handlerEvt As String = Nothing
                Dim handlerType As String = Nothing

                If modAccessEventCatalog.TryParseEventHandlerName(name, handlerObj, handlerEvt, handlerType) Then
                    currentMethodInfo.IsAccessEventHandler = True
                    currentMethodInfo.HandlerObject = handlerObj
                    currentMethodInfo.HandlerEvent = handlerEvt
                    currentMethodInfo.HandlerObjectType = handlerType

                    LogInfoLocal($" ↳ Access EventHandler detected: [{handlerObj}].[{handlerEvt}] ({handlerType})", 2)
                End If
            Catch ex As Exception
                LogError("TryParseEventHandlerName", ex)
            End Try

            For Each prm In paramsRaw
                prm.ParentMethod = currentMethodInfo
            Next

            currentMethodInfo.Parameters = paramsRaw
            'If tokenizedLine.Tokens.Count > 0 Then currentMethodInfo.TokenizedLine = tokenizedLine

            LogInfoLocal($" Parsed method declaration: {currentMethodInfo.Name} with {paramsRaw.Count} params.", 2)
        Catch ex As Exception
            LogError($" -> ({formOrModule.name}) ParseMetgodDelcaration", ex)
        End Try
    End Sub

    Private Function ParseVariable(L As String, isClassLike As Boolean, lineNumber As Integer, context As String, rxToUse As Regex, moduleName As String, rawLine As String, localLineNum As Integer, methodLineNum As Integer) As List(Of ParamInfo)
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
                        .LocalLineNumber = localLineNum,
                        .MethodLineNumber = methodLineNum,
                        .TokenString = pName,
                        .TokenType = TokenTypeEnum.variable_decl,  ' consistent
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

                            Dim tToken As New Token With {
                                .LineNumber = lineNumber,
                                .LocalLineNumber = localLineNum,
                                .MethodLineNumber = methodLineNum,
                                .TokenString = part,
                                .TokenType = TokenTypeEnum.unknown_type,
                                .Context = context,
                                .NextSymbol = nextSym,
                                .IsLeftSide = False
                            }
                            tokens.Add(tToken)
                            prevToken = tToken
                        Next

                        ' Salvez tipul în token-ul variabilei
                        nameToken.DataType = pType
                        nameToken.IsResolved = True
                        nameToken.ResolvedRef = prevToken
                    End If

                    ' Map tokeni la coloane în linia raw
                    MapTokensToColumns(tokens, rawLine, L)

                    Dim prm As New ParamInfo With {
                        .Name = pName,
                        .IsGlobal = isGlobal,
                        .IsPublic = rxToUse IsNot RegexCache.RxParams AndAlso (access = "Public" Or access = "" And Not isClassLike),
                        .Type = pType,
                        .Value = pValue,
                        .Tokens = tokens,
                        .IsWithEvents = isWithEvents
                    }

                    result.Add(prm)
                End If
            Next

            LogInfoLocal($" Parsed {result.Count} variables from line.", 2)

            Return result

        Catch ex As Exception
            LogError($" -> ({moduleName}:{context}) ParseVariabls", ex)
            Return New List(Of ParamInfo)
        End Try
    End Function

    Private Function PreParseCode(codeContent As String, moduleName As String) As List(Of (lineNum As Integer, text As String, rawLine As String))
        ' === Preproc: unește liniile cu underscore și numerotează ===
        Dim rawLines = codeContent.Split({vbCrLf, vbLf}, StringSplitOptions.None).ToList()
        Dim merged As New List(Of (lineNum As Integer, text As String, rawLine As String))
        Dim buffer As String = ""
        Dim ln As Integer = 1, startNum As Integer = 1
        Dim isContinued = False
        Dim rawIndex As Integer = 1

        Try
            For Each raw In rawLines
                LogInfoLocal($"({ln}) Raw Line: {raw}", 2)
                Dim L = raw.TrimEnd()
                rawIndex += 1
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
                        If cleaned <> "" Then merged.Add((startNum, cleaned, buffer))
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
                            merged.Add((ln, part & ":", L))
                        Else
                            Dim cleaned = CleanLineForShits(part.Trim())
                            If cleaned <> "" Then merged.Add((ln, cleaned, L))
                        End If
                    Next
                End If
                LogInfoLocal($"({ln}) Parsed Line: {L}", 2)
                ln += 1
            Next

            ' === rest dacă a rămas ceva în buffer ===
            If buffer <> "" Then
                For Each part In RegexCache.RxSplitColon.Split(buffer)
                    Dim cleaned = CleanLineForShits(part.Trim())
                    If cleaned <> "" Then merged.Add((startNum, cleaned, buffer))
                Next
            End If

            Return merged

        Catch ex As Exception
            LogError($" -> ({moduleName}) PreParseCode", ex)
            Return New List(Of (lineNum As Integer, text As String, rawLine As String))
        End Try
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

    ' ==============================================================
    ' MAP TOKEN COLUMNS - găsește coloana fiecărui token în linia raw
    ' ==============================================================
    Public Sub MapTokensToColumns(ByRef tokens As List(Of Token), rawLine As String, cleanLine As String)
        If tokens Is Nothing OrElse tokens.Count = 0 Then Return
        If String.IsNullOrEmpty(rawLine) AndAlso String.IsNullOrEmpty(cleanLine) Then Return

        Dim offsetRaw As Integer = 0
        Dim offsetClean As Integer = 0

        For Each token In tokens
            If String.IsNullOrEmpty(token.TokenString) Then Continue For

            ' Caută în RAW line pentru ColumnNumber
            If Not String.IsNullOrEmpty(rawLine) Then
                Dim foundPosRaw = rawLine.IndexOf(token.TokenString, offsetRaw, StringComparison.Ordinal)
                If foundPosRaw >= 0 Then
                    token.ColumnNumber = foundPosRaw + 1  ' 1-indexed
                    offsetRaw = foundPosRaw + token.TokenString.Length
                Else
                    token.ColumnNumber = 0
                End If
            Else
                token.ColumnNumber = 0
            End If

            ' Caută în CLEAN line pentru ColumnNumberLocal
            If Not String.IsNullOrEmpty(cleanLine) Then
                Dim foundPosClean = cleanLine.IndexOf(token.TokenString, offsetClean, StringComparison.Ordinal)
                If foundPosClean >= 0 Then
                    token.LocalColumnNumber = foundPosClean + 1  ' 1-indexed
                    token.MethodColumnNumber = foundPosClean + 1
                    offsetClean = foundPosClean + token.TokenString.Length
                Else
                    token.LocalColumnNumber = 0
                End If
            Else
                token.LocalColumnNumber = 0
            End If
        Next
    End Sub

    Private Sub BuildModuleSymbols(base As ModuleContainer)
        base.ModuleSymbols.Clear()

        ' === Variabile modul ===
        If base.Variables IsNot Nothing Then
            For Each v In base.Variables
                base.ModuleSymbols.Add(New SymbolEntry With {
                .Name = v.Name,
                .TypeName = v.Type,
                .DeclType = "Variable",
                .Scope = v.Scope,
                .SourceObject = v
            })
            Next
        End If

        ' === Constante modul ===
        If base.Constants IsNot Nothing Then
            For Each c In base.Constants
                base.ModuleSymbols.Add(New SymbolEntry With {
                .Name = c.Name,
                .TypeName = c.Type,
                .DeclType = "Const",
                .Scope = c.Scope,
                .SourceObject = c
            })
            Next
        End If

        ' === Funcții ===
        If base.Functions IsNot Nothing Then
            For Each fn In base.Functions.Values
                base.ModuleSymbols.Add(New SymbolEntry With {
                .Name = fn.Name,
                .TypeName = fn.ReturnType,
                .DeclType = "Function",
                .Scope = fn.MethodScope,
                .SourceObject = fn
            })
            Next
        End If

        ' === Proprietăți ===
        If base.Properties IsNot Nothing Then
            For Each p In base.Properties.Values
                base.ModuleSymbols.Add(New SymbolEntry With {
                .Name = p.Name,
                .TypeName = p.ReturnType,
                .DeclType = "Property",
                .Scope = p.MethodScope,
                .SourceObject = p
            })
            Next
        End If

        ' === Form Controls ===
        If TypeOf base Is FormReportContainer Then
            Dim f = DirectCast(base, FormReportContainer)
            If f.Controls IsNot Nothing Then
                For Each c In f.Controls
                    base.ModuleSymbols.Add(New SymbolEntry With {
                    .Name = c.Name,
                    .TypeName = c.Type,
                    .DeclType = "Control",
                    .Scope = "Private",
                    .SourceObject = c
                })
                Next
            End If
        End If

        ' NU adăugăm variabile locale în ModuleSymbols
    End Sub

    Private Sub AddToGlobalSymbols(mci As ModuleContainer)
        If mci.Type <> "Module" Then Exit Sub

        For Each sym In mci.ModuleSymbols
            If sym.Scope <> "Public" Then Continue For
            Dim key = $"{mci.Name}.{sym.Name}"
            If Not GlobalSymbols.ContainsKey(key) Then
                GlobalSymbols(key) = sym
            End If
        Next
    End Sub
End Module