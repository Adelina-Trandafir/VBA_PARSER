'PROJECT NAME: VBA_TOKENIZER
'FILE DESCRIPTION: Module pentru parsarea codului VBA în obiecte (metode, variabile, constante, declarații)
'PATH: VBA_TOKENIZER/modCodeParser.vb
Imports System.Collections.Concurrent
Imports System.Text.RegularExpressions
Imports System.Threading
Imports VBA_CORE
Imports VBA_CORE.Logger
Imports VBA_CORE.customTypes

Public Module modCodeParser
    Private ReadOnly pCurrentModule As New ThreadLocal(Of String)(Function() "")
    Private ReadOnly pCurrentType As New ThreadLocal(Of String)(Function() "")
    Private ReadOnly pCurrentMethod As New ThreadLocal(Of String)(Function() "")
    Private ReadOnly pCurrentEvent As New ThreadLocal(Of String)(Function() "")
    ' ==============================================================
    '                 BUILD SYMBOL DATABASE
    ' ==============================================================
    Public Sub ParseTextToObjects()
        ' === colecție surse: (Name, Type, Code, FCI)
        GlobalSources = New ConcurrentBag(Of (Name As String, Type As String, Code As String, FCI As Object))

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

            GlobalSources.OrderBy(Of String)(Function(s) s.Type AndAlso s.Name)

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

        Try
            MCI.Variables = New ParamInfoList
            MCI.Constants = New ParamInfoList
            MCI.Types = New EnumOrTypeList
            MCI.Enums = New EnumOrTypeList

        Catch ex As Exception
            LogError($"ParseSingleCodeBlock({currentModule}) - Init collections", ex)
        End Try

        'If currentModule = "clsDatabase" Then Stop

        Try
            Dim merged = PreParseCode(codeContent, currentModule)

            Dim currentClass As String = If(isClassLike, currentModule, "")
            Dim currentFunc As String = ""
            Dim beforeFirstFunc As Boolean = True
            Dim parsedLine As MethodLine
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

                parsedLine = New MethodLine With {
                    .LocalLineNumber = localLineNum,
                    .LineNumber = lineNumber,
                    .Content = L,
                    .OriginalContent = rawL,
                    .IsInEnumBlock = isInEnum,
                    .IsInTypeBlock = isInType
                }

                If withStack.Count > 0 Then
                    parsedLine.IsInWithBlock = True
                    parsedLine.WithBlockContext = withStack.Peek()
                    parsedLine.WithBlockDepth = withStack.Count + 1
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
                    parsedLine.IsTokenized = True

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
                    parsedLine.IsTokenized = True
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
                    parsedLine.IsTokenized = True
                    LogInfoLocal(" Conditional block (End): " & condBlock, 2)
                    'continue for

                ElseIf {"Private", "Public"}.Any(Function(p) L.TrimStart().StartsWith(p & " Declare", StringComparison.OrdinalIgnoreCase)) Then
                    ' Declarare functii API
                    ParseDeclares(L, MCI, lineNumber, condBlock)
                    AddLineToModuleAndMethod(MCI, currentMethodInfo, parsedLine)
                    parsedLine.IsDeclarationLine = True
                    parsedLine.IsTokenized = True

                    LogInfoLocal(" API: " & L, 2)
                    Continue For

                ElseIf {"Global Const", "Public Const", "Private Const", "Const"}.Any(Function(p) L.TrimStart.StartsWith(p, StringComparison.OrdinalIgnoreCase)) And Not isInMethod Then
                    ' Constante la nivel de modul
                    If MCI.Constants Is Nothing Then MCI.Constants = New ParamInfoList
                    MCI.Constants.AddRange(ParseVariable(L, isClassLike, lineNumber, currentModule, RegexCache.RxConsts, currentModule, rawL, localLineNum, 0))
                    parsedLine.IsDeclarationLine = True
                    parsedLine.IsTokenized = True
                    LogInfoLocal(" Constant: " & L, 2)
                    'continue for

                ElseIf {"Public Event", "Private Event", "Event"}.Any(Function(p) L.TrimStart.StartsWith(p, StringComparison.OrdinalIgnoreCase)) AndAlso currentMethodInfo Is Nothing Then
                    ' Declarare eveniment
                    parsedLine.IsEventLine = True
                    parsedLine.IsTokenized = True

                    ParseMethodDeclaration(L, MCI, currentMethodInfo, lineNumber, RegexCache.RxEvent, rawL, condBlock, localLineNum, 0)
                    If MCI.Events Is Nothing Then MCI.Events = New CustomDict(Of String, MethodInfo)("Events")
                    MCI.Events.Add(currentMethodInfo.Name, currentMethodInfo)
                    Dim evt As New EventSet With {
                        .EventName = currentMethodInfo.Name,
                        .HandlerString = MCI.Name,
                        .RefType = MCI.Type,
                        .RefObject = currentMethodInfo
                    }
                    MCI.EventSets.Add(currentMethodInfo.Name, evt)

                    AddLineToModuleAndMethod(MCI, currentMethodInfo, parsedLine)
                    currentMethodInfo = Nothing
                    LogInfoLocal(" Event: " & L, 2)

                    Continue For

                ElseIf RegexCache.RxMethodTester.IsMatch(L) Then
                    'inceput metodă/proprietate
                    methodLineNum = 0
                    ParseMethodDeclaration(L, MCI, currentMethodInfo, lineNumber, IIf(L.Contains("Property "), RegexCache.RxProperty, RegexCache.RxHeader), rawL, condBlock, localLineNum, methodLineNum)
                    pCurrentMethod.Value = currentMethodInfo.Name
                    parsedLine.StartsMethodBlock = True
                    LogInfoLocal(" Method/Property:" & L, 2)

                ElseIf {"Public Type", "Private Type", "Type"}.Any(Function(p) L.TrimStart.StartsWith(p, StringComparison.OrdinalIgnoreCase)) Then
                    ' Început Type
                    isInType = True
                    parsedLine.StartsTypeBlock = True
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
                    parsedLine.EndsTypeBlock = True
                    parsedLine.IsInTypeBlock = False

                    MCI.Types.Add(enumOrType)

                    LogInfoLocal($"Type '{enumOrType.Name}' parsed with {enumOrType.Members.Count} members.", 2)

                    enumOrType = Nothing
                    'continue for

                ElseIf {"Public Enum", "Private Enum", "Enum"}.Any(Function(p) L.TrimStart.StartsWith(p, StringComparison.OrdinalIgnoreCase)) Then
                    ' Început Enum
                    isInEnum = True
                    parsedLine.StartsEnumBlock = True
                    enumOrType = New EnumOrType With {
                        .Name = RegexCache.RxEnumTypeStart.Match(L).Groups(1).Value,
                        .Type = "Enum",
                        .IsPublic = Not L.TrimStart().StartsWith("Private", StringComparison.OrdinalIgnoreCase),
                        .IsGlobal = MCI.Type.Equals("Module", StringComparison.OrdinalIgnoreCase)
                    }

                    LogInfoLocal(" Enum: " & L, 2)

                ElseIf L.Trim = "End Enum" Then
                    ' Sfârșit Enum
                    parsedLine.EndsEnumBlock = True
                    parsedLine.IsInEnumBlock = False
                    isInEnum = False

                    MCI.Enums.Add(enumOrType)

                    LogInfoLocal($"Enum '{enumOrType.Name}' parsed with {enumOrType.Members.Count} members.", 2)

                    enumOrType = Nothing
                    'continue for

                ElseIf isInEnum Or isInType Then
                    ' Prelucrare in interior Type/Enum
                    parsedLine.IsInEnumBlock = isInEnum
                    parsedLine.IsInTypeBlock = isInType
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
                        If MCI.Variables Is Nothing Then MCI.Variables = New ParamInfoList
                        MCI.Variables.AddRange(ParseVariable(L, isClassLike, lineNumber, currentModule, RegexCache.RxGlobalVar, currentModule, rawL, localLineNum, methodLineNum))
                    Else
                        ' Variabile la nivel de modul
                        If currentMethodInfo.Variables Is Nothing Then currentMethodInfo.Variables = New ParamInfoList
                        currentMethodInfo.Variables.AddRange(ParseVariable(L, isClassLike, lineNumber, currentMethodInfo.Name, RegexCache.RxLocalVar, currentModule, rawL, localLineNum, localLineNum))
                    End If
                    parsedLine.IsDeclarationLine = True
                    LogInfoLocal(" Variable: " & L, 2)
                    'continue for

                ElseIf {"Function", "Sub", "Property"}.Any(Function(p) L.TrimStart().StartsWith("End " & p, StringComparison.OrdinalIgnoreCase)) Then
                    'sfârșit metodă/proprietate
                    If currentMethodInfo IsNot Nothing Then
                        currentMethodInfo.EndLine = lineNumber + 1
                        parsedLine.EndsMethodBlock = True

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

                        'parsedLine = New MethodLine With {.LineNumber = lineNumber, .Content = L}
                        AddLineToModuleAndMethod(MCI, currentMethodInfo, parsedLine)

                        'ResolveVariableScopes(currentMethodInfo, MCI)

                        LogInfoLocal(" Line parsed: " & L, 2)

                        ' === Finalizare metodă ===
                        Try
                            If MCI.Methods Is Nothing Then MCI.Methods = New CustomDict(Of String, MethodInfo)("Methods")

                            MCI.Methods.Add(fullMethodName, currentMethodInfo)

                            ' === Clasificare pe tip ===
                            If TypeOf currentMethodInfo Is FunctionInfo Then
                                If MCI.Functions Is Nothing Then MCI.Functions = New CustomDict(Of String, FunctionInfo)
                                MCI.Functions(fullMethodName) = DirectCast(currentMethodInfo, FunctionInfo)

                            ElseIf TypeOf currentMethodInfo Is PropertyInfo Then
                                If MCI.Properties Is Nothing Then MCI.Properties = New CustomDict(Of String, PropertyInfo)
                                MCI.Properties(fullMethodName) = DirectCast(currentMethodInfo, PropertyInfo)

                            ElseIf currentMethodInfo.Name.StartsWith("Event", StringComparison.OrdinalIgnoreCase) Then
                                If MCI.Events Is Nothing Then MCI.Events = New CustomDict(Of String, MethodInfo)
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

                    parsedLine.StartsWithBlock = True

                    withStack.Push(parsedLine)
                    LogInfoLocal($" With block START: '{withExpr}' (depth={withStack.Count})", 2)
                    'Continue For

                ElseIf L.TrimStart().StartsWith("End With", StringComparison.OrdinalIgnoreCase) Then
                    ' ===== Sfârșitul unui bloc With =====
                    parsedLine.EndsWithBlock = True
                    If withStack.Count > 0 Then
                        Dim closedWith = withStack.Pop()
                        LogInfoLocal($" With block END: '{closedWith}' (depth={withStack.Count})", 2)
                    Else
                        LogInfoLocal(" Warning: End With without matching With", 2)
                    End If
                    'Continue For

                ElseIf RegexCache.RxCodeLabel.IsMatch(L) AndAlso currentMethodInfo IsNot Nothing Then
                    ' ===== Linia este un LABEL =====
                    parsedLine.IsLabelLine = True
                    LogInfoLocal(" Label line: " & L, 2)
                Else

                End If

                ' === Procesare linie normală (în interior sau în afara unei metode)
                If currentMethodInfo IsNot Nothing Then
                    parsedLine.IsMethodLine = True
                    methodLineNum += 1
                End If

                AddLineToModuleAndMethod(MCI, currentMethodInfo, parsedLine)

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
        If formOrModule.methods IsNot Nothing Then
            For Each m In formOrModule.Methods?.Values
                For Each ml In m.MethodLines
                    totalTokens += ml.Tokens.Count
                Next
            Next
        End If

        Return $"Methods: {formOrModule.Methods?.Count}; " &
           $"Declares: {formOrModule.Declares?.Count}; " &
           $"Variables: {formOrModule.Variables?.Count}; " &
           $"Constants: {formOrModule.Constants?.Count}; "
    End Function

    Private Sub LogInfoLocal(ByVal msg As String, level As Integer, Optional force As Boolean = False)
        If pCurrentMethod.Value <> "" Then
            msg &= $"({pCurrentType.Value}) [{pCurrentModule.Value}].[{pCurrentMethod.Value}] : {msg}"
        Else
            msg &= $"({pCurrentType.Value}) [{pCurrentModule.Value}] : {msg}"
        End If

        LogInfo("[CODEPARSER]" & msg, level, force)
    End Sub

    ''' <summary>
    ''' Parsează o linie dintr-un bloc Enum sau Type și adaugă membrul în containerul părinte.
    ''' </summary>
    Private Sub ParseEnumTypeLine(L As String, ByRef enumOrType As EnumOrType, rxToUse As Regex)
        Try
            Dim raw = L.Trim()
            Dim isEnum As Boolean = (rxToUse Is RegexCache.RxEnumMember)
            Dim parentName As String = enumOrType.Name

            ' === 1️⃣ Extrage nume + tip / valoare ===
            Dim fieldName As String = ""
            Dim fieldType As String = ""
            Dim fieldValue As String = ""
            Dim m = rxToUse.Match(raw)

            If isEnum Then
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

            ' === 2️⃣ Creează calificativul complet ===
            Dim qualifiedName As String = $"{parentName}.{fieldName}"

            ' === 3️⃣ Creează tokenul membrului ===
            Dim memberTokenType As TokenTypeEnum = If(isEnum, TokenTypeEnum.enum_member_decl, TokenTypeEnum.type_member_decl)
            Dim memberToken As New Token With {
                .TokenString = fieldName,
                .TokenType = memberTokenType,
                .LineNumber = 0,
                .IsResolved = True,
                .ResolvedScope = If(isEnum, "enum_member_decl", "type_member_decl")
            }

            ' === 4️⃣ Creează obiectul ParamInfo ===
            Dim member As New ParamInfo With {
                .Name = fieldName,
                .Type = fieldType,
                .IsGlobal = enumOrType.IsGlobal,
                .IsPublic = enumOrType.IsPublic,
                .QualifiedName = qualifiedName,
                .Tokens = New List(Of Token) From {memberToken}
            }

            If isEnum Then member.Value = fieldValue
            enumOrType.Members.Add(member)

            LogInfoLocal($"Parsed {If(isEnum, "Enum", "Type")} member: {qualifiedName} ({fieldType})", 2)

        Catch ex As Exception
            LogError($"ParseEnumTypeLine({enumOrType?.Name})", ex)
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
        Dim isFunction As Boolean = False

        Try
            If h.Groups(1).Success Then isPublic = h.Groups(1).Value = "Public"
            If h.Groups(3).Success Then kind = h.Groups(3).Value
            If h.Groups(4).Success Then name = h.Groups(4).Value
            If h.Groups(5).Success Then fLib = h.Groups(5).Value
            If h.Groups(6).Success Then fAlias = h.Groups(6).Value
            If h.Groups(7).Success Then inner = h.Groups(7).Value.Trim()
            If h.Groups(8).Success Then returnType = h.Groups(8).Value

            paramsRaw = New ParamInfoList
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
                isFunction = True
                currentMethodInfo = New FunctionInfo With {
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
                .MethodLines = New MethodLineList
                .Parameters = paramsRaw
                .isfunction = isFunction
            End With

            For Each prm In paramsRaw
                prm.ParentMethod = currentMethodInfo
            Next

            currentMethodInfo.Parameters = paramsRaw

            If formOrModule.declares Is Nothing Then formOrModule.declares = New CustomDict(Of String, MethodInfo)("Declares")
            formOrModule.Declares.add(fullName, currentMethodInfo)

            LogInfoLocal($" Parsed declare: {currentMethodInfo.Name} with {paramsRaw.Count} params.", 2)
        Catch ex As Exception
            LogError($" -> ({formOrModule.name}) ParseDeclare", ex)
            'Stop
        End Try
    End Sub

    Private Sub ParseMethodDeclaration(L As String, formOrModule As Object, ByRef currentMethodInfo As MethodInfo, lineNumber As Long, rxToUse As Regex, rawLine As String, condBlock As String, localLineNumber As Integer, methodLineNumber As Integer)
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

            paramsRaw = New ParamInfoList
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
                        .IsGlobal = False,
                        .Tokens = New List(Of Token)
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
                .MethodLines = New MethodLineList
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
            'If parsedLine.Tokens.Count > 0 Then currentMethodInfo.ParsedLine = parsedLine

            LogInfoLocal($" Parsed method declaration: {currentMethodInfo.Name} with {paramsRaw.Count} params.", 2)
        Catch ex As Exception
            LogError($" -> ({formOrModule.name}) ParseMetgodDelcaration", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Parsează o linie care conține una sau mai multe declarații de variabile sau constante
    ''' (Dim / Const / Public / Private / Global) și returnează lista de ParamInfo.
    ''' </summary>
    Private Function ParseVariable(L As String, isClassLike As Boolean, lineNumber As Integer, context As String, rxToUse As Regex, moduleName As String, rawLine As String, localLineNum As Integer, methodLineNum As Integer) As List(Of ParamInfo)
        Dim result As New List(Of ParamInfo)

        Try
            Dim matches = rxToUse.Matches(L)
            Dim isGlobal As Boolean = False

            For Each m As Match In matches
                If Not m.Success Then Continue For

                ' === 1️⃣ Extrage valorile din regex ===
                Dim access As String = ""
                Dim isWithEvents As Boolean = False
                Dim pName As String = ""
                Dim pType As String = ""
                Dim pValue As String = ""

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

                ' === 2️⃣ Construiește lista de tokeni ===
                Dim tokens As New List(Of Token)

                ' Numele variabilei
                Dim nameToken As New Token With {
                    .LineNumber = lineNumber,
                    .LocalLineNumber = localLineNum,
                    .MethodLineNumber = methodLineNum,
                    .TokenString = pName,
                    .TokenType = If(rxToUse Is RegexCache.RxConsts, TokenTypeEnum.constant_decl, TokenTypeEnum.variable_decl),
                    .Context = context,
                    .Parent = Nothing,
                    .NextSymbol = If(String.IsNullOrEmpty(pType), "", "As"),
                    .IsLeftSide = True,
                    .IsResolved = True
                }

                tokens.Add(nameToken)

                ' Tipul de date
                If Not String.IsNullOrEmpty(pType) Then
                    Dim typeParts = pType.Split("."c)
                    Dim prevToken As Token = nameToken

                    For i = 0 To typeParts.Length - 1
                        Dim part = typeParts(i).Trim()
                        Dim nextSym As String = If(i < typeParts.Length - 1, ".", "")
                        If part = "" Then Continue For

                        Dim tToken As New Token With {
                            .LineNumber = lineNumber,
                            .LocalLineNumber = localLineNum,
                            .MethodLineNumber = methodLineNum,
                            .TokenString = part,
                            .TokenType = TokenTypeEnum.type_ref,
                            .Context = context,
                            .NextSymbol = nextSym,
                            .IsLeftSide = False,
                            .IsResolved = True
                        }

                        tokens.Add(tToken)
                        prevToken = tToken
                    Next

                    ' leagă tipul în token-ul principal
                    nameToken.DataType = pType
                    nameToken.ResolvedRef = tokens.Last()
                Else
                    nameToken.DataType = "Variant"
                End If

                ' === 3️⃣ Mapează tokenii pe coloane ===
                MapTokensToColumns(tokens, rawLine, L)

                ' === 4️⃣ Creează ParamInfo ===
                Dim prm As New ParamInfo With {
                    .Name = pName,
                    .IsGlobal = isGlobal,
                    .IsPublic = rxToUse IsNot RegexCache.RxParams AndAlso (access = "Public" OrElse (access = "" AndAlso Not isClassLike)),
                    .Type = pType,
                    .Value = pValue,
                    .Tokens = tokens,
                    .IsWithEvents = isWithEvents
                }

                result.Add(prm)
            Next

            LogInfoLocal($"Parsed {result.Count} variables from line.", 2)
            Return result

        Catch ex As Exception
            LogError($"ParseVariable({moduleName}:{context})", ex)
            Return New List(Of ParamInfo)
        End Try
    End Function

    ''' <summary>
    ''' Preprocesează codul sursă VBA pentru a combina liniile continuate (cu underscore) și a separa
    ''' instrucțiunile multiple de pe aceeași linie (delimitate prin „:”).
    ''' Detectează și etichetele de cod (Label:) și le păstrează ca linii distincte fără a le despărți.
    ''' </summary>
    ''' <param name="codeContent">Conținutul complet al modulului VBA de analizat (text brut).</param>
    ''' <param name="moduleName">Numele modulului curent, utilizat pentru logare și raportare erori.</param>
    ''' <returns>
    ''' O listă de tuple (lineNum, text, rawLine):
    '''   • <c>lineNum</c> — numărul liniei originale din fișier;
    '''   • <c>text</c> — linia procesată (curățată și combinată);
    '''   • <c>rawLine</c> — textul brut original corespunzător.
    ''' </returns>
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

                ' === ignoră comentarii, linii goale, Option ===
                If L.TrimStart().StartsWith("'@@") Then
                    Continue For
                ElseIf L = "" OrElse L.TrimStart().StartsWith("'") Then
                    ln += 1 : Continue For
                ElseIf L.TrimStart().StartsWith("Option ", StringComparison.OrdinalIgnoreCase) Then
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
                    Dim trimmed = L.Trim()

                    ' 🔹 test special pentru label (doar un cuvânt urmat de ":")
                    If RegexCache.RxCodeLabel.IsMatch(trimmed) Then
                        merged.Add((ln, trimmed, L))

                    Else
                        ' 🔹 altfel, o linie normală care poate conține mai multe instrucțiuni separate de ":"
                        For Each part In RegexCache.RxSplitColon.Split(L)
                            Dim p = part.Trim()

                            ' dacă e un singur cuvânt fără spații, poate fi destinație GoTo
                            If p.Split(" "c).Length = 1 AndAlso Not String.IsNullOrEmpty(p) Then
                                merged.Add((ln, p, L))
                            Else
                                Dim cleaned = CleanLineForShits(p)
                                If cleaned <> "" Then merged.Add((ln, cleaned, L))
                            End If
                        Next
                    End If
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
        If base.ModuleSymbols Is Nothing Then base.ModuleSymbols = New CustomDict(Of String, SymbolEntry)
        base.ModuleSymbols.Clear()

        ' === Variabile modul ===
        If base.Variables IsNot Nothing Then
            For Each v In base.Variables
                base.ModuleSymbols(v.Name) = New SymbolEntry With {
                    .Name = v.Name,
                    .TypeName = v.Type,
                    .DeclType = TokenTypeEnum.variable_decl,
                    .Scope = v.Scope,
                    .SourceObject = v
                }
            Next
        End If

        ' === Constante modul ===
        If base.Constants IsNot Nothing Then
            For Each c In base.Constants
                base.ModuleSymbols(c.Name) = New SymbolEntry With {
                    .Name = c.Name,
                    .TypeName = c.Type,
                    .DeclType = TokenTypeEnum.constant_decl,
                    .Scope = c.Scope,
                    .SourceObject = c
                }
            Next
        End If

        ' === Funcții ===
        If base.Functions IsNot Nothing Then
            For Each fn In base.Functions.Values
                base.ModuleSymbols(fn.Name) = New SymbolEntry With {
                    .Name = fn.Name,
                    .TypeName = fn.ReturnType,
                    .DeclType = TokenTypeEnum.function_decl,
                    .Scope = fn.MethodScope,
                    .SourceObject = fn
                }
            Next
        End If

        ' === Proprietăți ===
        If base.Properties IsNot Nothing Then
            For Each p In base.Properties.Values
                base.ModuleSymbols(p.Name) = New SymbolEntry With {
                    .Name = p.Name,
                    .TypeName = p.ReturnType,
                    .DeclType = If(p.PropertyType = "Get", TokenTypeEnum.property_get_decl,
                                  If(p.PropertyType = "Let", TokenTypeEnum.property_let_decl, TokenTypeEnum.property_set_decl)),
                    .Scope = p.MethodScope,
                    .SourceObject = p
                }
            Next
        End If

        ' === Form Controls ===
        If TypeOf base Is FormReportContainer Then
            Dim f = DirectCast(base, FormReportContainer)
            For Each c In f.Controls
                base.ModuleSymbols(c.Name) = New SymbolEntry With {
                    .Name = c.Name,
                    .TypeName = c.Type,
                    .DeclType = TokenTypeEnum.access_control,
                    .Scope = "Private",
                    .SourceObject = c
                }
            Next
        End If
    End Sub

    Private Sub AddToGlobalSymbols(mci As ModuleContainer)
        If mci Is Nothing OrElse mci.Type <> "Module" Then Exit Sub
        If mci.ModuleSymbols Is Nothing OrElse mci.ModuleSymbols.Count = 0 Then Exit Sub

        For Each kvp In mci.ModuleSymbols
            Dim sym = kvp.Value
            If sym.Scope <> "Public" Then Continue For

            Dim key = $"{mci.Name}.{sym.Name}"
            If Not GlobalSymbols.ContainsKey(key) Then
                GlobalSymbols(key) = sym
            End If
        Next
    End Sub

    ''' <summary>
    ''' Adaugă o linie de cod atât în lista modulului (<see cref="ModuleContainer.Lines"/>),
    ''' cât și în lista metodei curente (<see cref="MethodInfo.MethodLines"/>), 
    ''' păstrând aceeași instanță de <see cref="MethodLine"/> pentru consistență completă.
    ''' </summary>
    ''' <param name="MCI">Containerul modulului curent (Form, Report, Class, Module).</param>
    ''' <param name="currentMethodInfo">Metoda curentă, dacă există (altfel <c>Nothing</c>).</param>
    ''' <param name="line">Linia care trebuie adăugată.</param>
    Private Sub AddLineToModuleAndMethod(ByRef MCI As ModuleContainer, ByRef currentMethodInfo As MethodInfo, ByRef line As MethodLine)
        Try
            If line.Tokens?.Count > 0 Then Stop

            ' === Adaugă în metoda curentă (dacă există)
            If currentMethodInfo IsNot Nothing Then
                currentMethodInfo.MethodLines.Add(line)
                If MCI.Lines Is Nothing Then MCI.Lines = New MethodLineList

                ' Asigură-te că MCI folosește exact aceeași instanță
                MCI.Lines.Add(currentMethodInfo.MethodLines.Last())
            Else
                If MCI.Lines Is Nothing Then MCI.Lines = New MethodLineList
                ' Dacă nu există metodă curentă, adaugă direct
                MCI.Lines.Add(line)
            End If
        Catch ex As Exception
            LogError($"AddLineToModuleAndMethod({MCI?.Name})", ex)
        End Try
    End Sub

End Module