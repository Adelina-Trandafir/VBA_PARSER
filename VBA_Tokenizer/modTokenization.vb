'PROJECT NAME: VBA_TOKENIZER
'FILE DESCRIPTION: Module pentru tokenizarea codului VBA
'PATH: VBA_TOKENIZER/modTokenization.vb

Imports System.Text
Imports System.Text.RegularExpressions
Imports VBA_CORE
Imports VBA_CORE.customTypes
Imports VBA_CORE.modAccessPropertyCatalog

Public Module Tokenizer
    Private ReadOnly pCurrentType As New Threading.ThreadLocal(Of String)(Function() "")
    Private ReadOnly pCurrentModule As New Threading.ThreadLocal(Of String)(Function() "")
    Private ReadOnly pCurrentMethod As New Threading.ThreadLocal(Of String)(Function() "")

    ''' <summary>
    ''' Tokenizează toate modulele din GlobalSources (Forms, Reports, Classes, Modules).
    ''' Rulează paralel sau secvențial în funcție de Global_UseParallel.
    ''' </summary>
    Public Sub TokenizeAllModules()
        If GlobalSources Is Nothing OrElse GlobalSources.Count = 0 Then Exit Sub

        Dim opts As New ParallelOptions With {.MaxDegreeOfParallelism = Environment.ProcessorCount}

        Dim runner As Action =
        Sub()
            If modGlobals.Global_UseParallel Then
                Parallel.ForEach(GlobalSources, opts,
                    Sub(src)
                        Try
                            TokenizeModule(src.FCI)
                        Catch ex As Exception
                            LogError($"TokenizeModule({src.Name})", ex)
                        End Try
                    End Sub)
            Else
                For Each src In GlobalSources
                    Try
                        TokenizeModule(src.FCI)
                    Catch ex As Exception
                        LogError($"TokenizeModule({src.Name})", ex)
                    End Try
                Next
            End If
        End Sub

        runner()
    End Sub

    ''' <summary>
    ''' Tokenizează toate metodele și liniile dintr-un modul (Form/Class/Module).
    ''' Rulează paralel sau secvențial în funcție de Global_UseParallel.
    ''' </summary>
    Public Sub TokenizeModule(modObj As ModuleContainer)
        If modObj Is Nothing OrElse modObj.Methods?.Count = 0 Then Exit Sub

        Try
            pCurrentModule.Value = modObj.Name
            pCurrentType.Value = modObj.Type

            Dim opts As New ParallelOptions With {.MaxDegreeOfParallelism = Environment.ProcessorCount}
            'If modObj.Name = "clsDatabase" Then Stop

            LogInfoLocal($"Tokenizing module: {pCurrentModule.Value} ({pCurrentType.Value})", 1)

            ' === Tokenizează liniile la nivel de modul (în afara metodelor) ===
            If modObj.Lines IsNot Nothing Then
                LogInfoLocal($"Tokenizing {modObj.Lines.Count} module-level lines in module: {pCurrentModule.Value} ({pCurrentType.Value})", 2)
                For Each ml In modObj.Lines.Where(Function(t) Not t.IsMethodLine)
                    If Not ml.StartsMethodBlock AndAlso Not ml.EndsMethodBlock AndAlso Not ml.IsMethodLine Then
                        Dim toks = GetTokensFromLine(Nothing, modObj, ml)
                        ml.Tokens = toks
                    End If
                Next
            End If

            PrepareWorkingLines(modObj)
            'PreParseTokens(modObj)
            ResolveAccessEventHandlers(modObj)

            If modObj.Methods IsNot Nothing Then
                LogInfoLocal($"Tokenizing {modObj.Methods.Count} methods in module: {pCurrentModule.Value} ({pCurrentType.Value})", 2)

                For Each m As MethodInfo In modObj.Methods.Values
                    pCurrentMethod.Value = m.Name

                    For Each ml As MethodLine In m.MethodLines.Where(Function(l) l.IsLabelLine)
                        Dim toks = New TokenList From {GetTokensFromLabelLine(m, ml)}
                        If m.CodeLabels Is Nothing Then m.CodeLabels = New TokenList
                        m.CodeLabels.AddRange(toks)
                    Next
                Next
            End If

            ' === Tokenize fiecare metodă ===
            Dim methodAction As Action(Of MethodInfo) =
            Sub(m)
                ' Stivă locală pentru blocuri WITH (poate fi extinsă)
                Dim withStack As New Stack(Of MethodLine)
                Dim toks As TokenList = Nothing

                Try
                    For Each ml In m.MethodLines.Where(Function(l) Not l.IsLabelLine)
                        If ml.IsTokenized Then Continue For
                        LogInfoLocal($"Tokenizing line {ml.LineNumber} in method: {pCurrentMethod.Value} (module: {pCurrentModule.Value})", 3)

                        ' === gestionează blocurile WITH ===
                        If ml.StartsWithBlock Then
                            withStack.Push(ml)
                        ElseIf ml.EndsWithBlock Then
                            If withStack.Count > 0 Then withStack.Pop()
                        ElseIf ml.IsInWithBlock AndAlso withStack.Count > 0 Then
                            ml.WithBlockContext = withStack.Peek()
                        End If

                        Try

                            toks = GetTokensFromLine(m, modObj, ml)

                            Try
                                ml.Tokens.AddRange(toks)

                                If m.Tokens Is Nothing Then m.Tokens = New TokenList
                                m.Tokens.AddRange(toks)

                            Catch ex As Exception
                                LogError("TokenizeModule.Method.Line.AddTokens", ex)
                            End Try

                        Catch ex As Exception
                            LogError("TokenizeModule.Method.Line", ex)
                        End Try
                    Next
                Catch ex As Exception
                    LogError("TokenizeModule.Method", ex)
                End Try
            End Sub

            If modGlobals.Global_UseParallel Then
                Parallel.ForEach(modObj.Methods.Values, opts, methodAction)
            Else
                If modObj.Methods IsNot Nothing Then
                    For Each m In modObj.Methods.Values
                        methodAction(m)
                    Next
                End If
            End If

        Catch ex As Exception
            LogError("TokenizeModule", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Curăță liniile sursă și le pregătește pentru tokenizare.
    ''' Elimină comentariile (dacă există deja filtrate) și înlocuiește toate
    ''' stringurile (între ghilimele) cu markere de forma &lt;STR_start_len_end&gt;.
    ''' Rezultatul este salvat în .WorkingLines pentru modul și pentru fiecare metodă.
    ''' </summary>
    ''' <param name="currentModule">Containerul modulului curent.</param>
    Private Sub PrepareWorkingLines(ByRef currentModule As ModuleContainer)
        Try
            If currentModule.Lines Is Nothing OrElse currentModule.Lines.Count = 0 Then Exit Sub

            ' === 1️⃣ Procesează codul la nivel de modul (în afara metodelor) ===
            Dim processedLines As New MethodLineList
            For Each line In currentModule.Lines.Where(Function(ln) Not ln.StartsMethodBlock AndAlso Not ln.IsMethodLine AndAlso Not ln.EndsMethodBlock)
                Dim workingLine As New MethodLine
                workingLine = line
                workingLine.Content = ReplaceStringsWithMarkers(line.Content)
                PreParseWorkingLine(workingLine, line)
                processedLines.Add(workingLine)
            Next
            currentModule.WorkingLines = processedLines

            ' === 2️⃣ Procesează fiecare metodă ===
            For Each m In currentModule.Methods.Values
                processedLines = New MethodLineList
                If m.MethodLines Is Nothing OrElse m.MethodLines.Count = 0 Then Continue For

                For Each line In m.MethodLines
                    Dim workingLine As New MethodLine
                    workingLine = line
                    workingLine.Content = ReplaceStringsWithMarkers(line.Content)
                    PreParseWorkingLine(workingLine, line)
                    processedLines.Add(workingLine)
                Next
                m.WorkingLines = processedLines
            Next

        Catch ex As Exception
            LogError("PrepareWorkingLines", ex)
        End Try
    End Sub

    Friend Sub GetVariables(ByRef ln As MethodLine)

    End Sub
    ''' <summary>
    ''' Înlocuiește toate stringurile dintre ghilimele cu markere de forma &lt;STR_start_len_end&gt;.
    ''' Utilizează expresia regulată globală RxLineCleaner pentru performanță.
    ''' </summary>
    Private Function ReplaceStringsWithMarkers(line As String) As String
        Return RegexCache.RxLineCleaner.Replace(line,
        Function(m)
            Dim startIdx As Integer = m.Index
            Dim length As Integer = m.Length
            Dim endIdx As Integer = startIdx + length - 1
            Return $"<STR_{startIdx}_{length}_{endIdx}>"
        End Function)
    End Function

    Friend Sub PreParseWorkingLine(ByRef wln As MethodLine, ln As MethodLine)
        For Each m As Match In RegexCache.RxPreTokenizer.Matches(wln.Content)
            If m.Success AndAlso Not String.IsNullOrWhiteSpace(m.Value) Then
                Dim tok As New Token With {
                                            .LineNumber = ln.LineNumber,
                                            .LocalLineNumber = ln.LocalLineNumber,
                                            .TokenString = m.Value.Trim(),
                                            .SourceLine = ln,
                                            .TokenType = TokenTypeEnum.initial
                                            }
                ln.PreTokens.Add(tok)
            End If
        Next
    End Sub

    '    Public Shared ReadOnly RxPreTokenizer As New Regex("(?>\<[^>]*\>)|(\s*[A-Za-z_]\w*(?:[.!][A-Za-z_]\w*)*(?:\s+As\s+[A-Za-z_]\w*)?)", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    ''' <summary>
    ''' Prelucreaza liniile pentru a extrage tokeni preliminari.
    ''' </summary>
    ''' <param name="currentModule"></param>
    Friend Sub PreParseTokens(currentModule As ModuleContainer)
        Dim tl As New TokenList

        Try
            For Each l In currentModule.WorkingLines.Where(Function(ln) Not ln.StartsMethodBlock AndAlso Not ln.IsMethodLine AndAlso Not ln.EndsMethodBlock)
                For Each m As Match In RegexCache.RxPreTokenizer.Matches(l.Content)
                    If m.Success AndAlso Not String.IsNullOrWhiteSpace(m.Value) Then
                        Dim tok As New Token With {
                                                    .LineNumber = l.LineNumber,
                                                    .LocalLineNumber = l.LocalLineNumber,
                                                    .TokenString = m.Value.Trim(),
                                                    .SourceLine = l,
                                                    .WithBlockDepth = l.WithBlockDepth,
                                                    .TokenType = TokenTypeEnum.initial
                                                    }
                        tl.Add(tok)
                    End If
                Next
                l.PreTokens = tl
            Next
        Catch ex As Exception
            LogError("PreParseTokens", ex)
        End Try

        Try
            For Each mtd In currentModule.Methods.Values '.SelectMany(Function(m) m.WorkingLines) '.Where(Function(ln) Not ln.IsTokenized)
                tl = New TokenList
                For Each ml In mtd.WorkingLines
                    For Each m As Match In RegexCache.RxPreTokenizer.Matches(ml.Content)
                        If m.Success AndAlso Not String.IsNullOrWhiteSpace(m.Value) Then
                            Dim tok = New Token With {
                                                    .LineNumber = ml.LineNumber,
                                                    .LocalLineNumber = ml.LocalLineNumber,
                                                    .TokenString = m.Value.Trim(),
                                                    .SourceLine = ml,
                                                    .WithBlockDepth = ml.WithBlockDepth,
                                                    .TokenType = TokenTypeEnum.initial
                                                    }
                            tl.Add(tok)
                        End If
                    Next
                    ml.PreTokens = tl
                Next
            Next
        Catch ex As Exception
            LogError("PreParseTokens.MethodLines", ex)
        End Try
    End Sub

    ' ==============================================================
    ' 🔹 Extrage tokeni dintr-o linie curățată
    ' ==============================================================
    Private Function GetTokensFromLine(ByRef currentMethod As MethodInfo, ByRef currentModule As ModuleContainer, ByRef line As MethodLine, Optional methodContext As String = "") As TokenList
        Dim tokens As New TokenList
        Dim prevOp As String = Nothing
        Dim asContextTarget As Token = Nothing

        'If line.Content = "Private Sub Form_Load()" Then Stop
        If currentMethod IsNot Nothing Then
            pCurrentMethod.Value = currentMethod.Name
        Else
            pCurrentMethod.Value = ""
        End If

        Try
            Dim cleanLine As String = line.Content
            If String.IsNullOrWhiteSpace(cleanLine) Then Return tokens

            Dim withContext As MethodLine = If(line.IsInWithBlock, line.WithBlockContext, Nothing)
            Dim startsWithDot As Boolean = False

            For Each m As Match In RegexCache.RxTokenizer.Matches(cleanLine)
                Dim fullExpression = m.Groups(1).Value
                Dim op = m.Groups(2).Value
                If String.IsNullOrWhiteSpace(fullExpression) Then Continue For
                If Char.IsDigit(fullExpression(0)) Then Continue For

                Dim parts() As String = fullExpression.Split("."c)
                startsWithDot = fullExpression.StartsWith(".")
                If startsWithDot Then
                    parts = parts.Where(Function(p) Not String.IsNullOrEmpty(p)).ToArray()
                End If

                Dim chainParent As Token = Nothing

                For i As Integer = 0 To parts.Length - 1
                    Dim partName = parts(i).Replace("()", "").Trim()
                    If partName = "" Then Continue For

                    Dim tok As New Token With {
                            .LineNumber = line.LineNumber,
                            .LocalLineNumber = line.LocalLineNumber,
                            .MethodLineNumber = i + 1,
                            .TokenString = partName,
                            .NextSymbol = If(i < parts.Length - 1, ".", op),
                            .Context = methodContext,
                            .SourceLine = line,
                            .WithBlockDepth = line.WithBlockDepth,
                            .TokenType = TokenTypeEnum.unresolved
                        }

                    LogInfoLocal($"Identified token: '{tok.TokenString}' (next: '{tok.NextSymbol}')", 3)

                    ' === context "As" activ (următorul token e tipul declarat)
                    If asContextTarget IsNot Nothing Then
                        asContextTarget.DataType = tok.TokenString
                        asContextTarget.TokenType = TokenTypeEnum.type_ref
                        asContextTarget = Nothing
                        tok.TokenType = TokenTypeEnum.unresolved
                    End If

                    ' === clasificare built-in / Access / keyword ===
                    Dim category = modIgnoreLists.GetIgnoreCategory(partName)
                    Select Case category
                        Case "VBA Type" : tok.TokenType = TokenTypeEnum.builtin_type : tok.IsBuiltIn = True
                        Case "VBA Function" : tok.TokenType = TokenTypeEnum.builtin_function : tok.IsBuiltIn = True
                        Case "VBA Object" : tok.TokenType = TokenTypeEnum.builtin_object : tok.IsBuiltIn = True
                        Case "VBA Constant" : tok.TokenType = TokenTypeEnum.builtin_constant : tok.IsBuiltIn = True
                        Case "Access Object" : tok.TokenType = TokenTypeEnum.access_object : tok.IsBuiltIn = True
                        Case "Access Function" : tok.TokenType = TokenTypeEnum.access_function : tok.IsBuiltIn = True
                        Case "VBA Keyword" : tok.TokenType = TokenTypeEnum.vba_keyword : tok.IsBuiltIn = True
                        Case Else : tok.TokenType = TokenTypeEnum.unresolved
                    End Select

                    ' === declarații Dim / Const ===
                    If line.IsDeclarationLine AndAlso tok.TokenType = TokenTypeEnum.unresolved Then
                        Dim isConstDecl = line.Content.TrimStart().StartsWith("Const ", StringComparison.OrdinalIgnoreCase)
                        If tok.NextSymbol.Equals("As", StringComparison.OrdinalIgnoreCase) Then
                            tok.TokenType = If(isConstDecl, TokenTypeEnum.constant_decl, TokenTypeEnum.variable_decl)
                            tok.ResolvedScope = If(currentMethod Is Nothing, "module_decl", "local_decl")
                            tok.IsResolved = True
                            tok.DataType = ""
                            asContextTarget = tok
                        ElseIf tok.NextSymbol = "," OrElse tok.NextSymbol = "" Then
                            tok.TokenType = If(isConstDecl, TokenTypeEnum.constant_decl, TokenTypeEnum.variable_decl)
                            tok.ResolvedScope = If(currentMethod Is Nothing, "module_decl", "local_decl")
                            tok.IsResolved = False
                            tok.DataType = "Variant"
                        End If
                        LogInfoLocal($"Declaration token: '{tok.TokenString}' as '{tok.TokenType}'", 3)
                    End If

                    ' === token special "Me" ===
                    If tok.TokenString.Equals("Me", StringComparison.OrdinalIgnoreCase) Then
                        tok.TokenType = TokenTypeEnum.self
                        tok.IsBuiltIn = currentModule.Type <> "Class"
                    End If

                    If withContext IsNot Nothing AndAlso startsWithDot Then
                        tok.IsWithMember = True
                        tok.WithBlockContext = withContext
                        tok.ResolvedScope = "with_context"
                        tok.TokenType = TokenTypeEnum.variable_member
                        tok.QualifiedName = BuildQualifiedWithName(withContext, tok)
                        ResolveMemberFromContext(tok, withContext, currentModule)
                    End If

                    If chainParent IsNot Nothing Then
                        tok.Parent = chainParent
                        ResolveMemberFromContext(tok, chainParent, currentModule)
                    End If

                    If tok.IsResolved Then
                        tokens.Add(tok)
                        chainParent = tok
                        Continue For
                    End If

                    ' === tokeni rezolvabili ===
                    If Not tok.IsBuiltIn Then
                        tok.IsResolved = False
                        tokens.Add(tok)
                        chainParent = tok

                        If Not tok.NextSymbol.ToUpper.Contains("AS") Then
                            ResolveTokenScope(tok, currentMethod, currentModule)

                            'If Not tok.IsResolved Then
                            '    If line.StartsMethodBlock AndAlso currentMethod IsNot Nothing Then
                            '        TokenizerResolves.ResolveAccessEventHandler(tok, currentMethod, currentModule)
                            '    End If
                            'End If
                        Else
                            asContextTarget = tok
                        End If

                        LogInfoLocal($"Resolved token: '{tok.TokenString}' as '{tok.TokenType}' (scope: '{tok.ResolvedScope}')", 3)
                    Else
                        tok.IsResolved = True
                        tokens.Add(tok)

                        If tok.TokenType <> TokenTypeEnum.vba_keyword AndAlso tok.TokenType <> TokenTypeEnum.builtin_constant Then chainParent = tok
                    End If
                Next

                prevOp = op
            Next

            ' === mapare poziții + rezoluție recursivă ===
            'If line.Tokens Is Nothing Then line.Tokens = New List(Of Token)
            'line.Tokens.AddRange(tokens)
            MapTokensToColumns(tokens, line.OriginalContent, cleanLine)

            For Each tok In tokens
                If tok.TokenType = TokenTypeEnum.unresolved AndAlso Not tok.IsBuiltIn Then
                    ResolveTokenRecursive(tok, currentMethod, currentModule)
                End If
            Next

            Return tokens

        Catch ex As Exception
            LogError("GetTokensFromLine", ex)
            Return New List(Of Token)
        End Try
    End Function

    ''' <summary>
    ''' Adaugă tokenii care definesc variabile, parametri sau constante în contextul metodei curente.
    ''' </summary>
    Private Sub AddTokenToMethodContext(ByRef tok As Token, currentMethod As MethodInfo)
        If currentMethod Is Nothing Then Exit Sub
        Dim curTok As Token = tok
        Try
            Select Case tok.TokenType
                Case TokenTypeEnum.variable_decl
                    ' Dacă variabila nu există deja, o adăugăm
                    If Not currentMethod.Variables.Any(Function(v) v.Name.Equals(curTok.TokenString, StringComparison.OrdinalIgnoreCase)) Then
                        currentMethod.Variables.Add(New ParamInfo With {
                        .Name = tok.TokenString,
                        .DataType = tok.DataType,
                        .DeclaredInLine = tok.LineNumber
                    })
                    End If

                Case TokenTypeEnum.parameter_decl
                    If Not currentMethod.Parameters.Any(Function(p) p.Name.Equals(curTok.TokenString, StringComparison.OrdinalIgnoreCase)) Then
                        currentMethod.Parameters.Add(New ParamInfo With {
                        .Name = tok.TokenString,
                        .DataType = tok.DataType,
                        .DeclaredInLine = tok.LineNumber
                    })
                    End If

                Case TokenTypeEnum.constant_decl
                    If Not currentMethod.Constants.Any(Function(c) c.Name.Equals(curTok.TokenString, StringComparison.OrdinalIgnoreCase)) Then
                        currentMethod.Constants.Add(New ParamInfo With {
                        .Name = tok.TokenString,
                        .DataType = tok.DataType,
                        .DeclaredInLine = tok.LineNumber
                    })
                    End If
            End Select
        Catch ex As Exception
            LogError("AddTokenToMethodContext", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Adaugă tokenii declarați la nivel de modul (variabile, constante, declarații API, metode publice).
    ''' </summary>
    Private Sub AddTokenToModuleContext(ByRef tok As Token, ByRef currentModule As ModuleContainer)
        If currentModule Is Nothing Then Exit Sub
        Dim locTok As Token = tok

        Select Case tok.TokenType
            Case TokenTypeEnum.variable_decl
                If Not currentModule.Variables.Any(Function(v) v.Name.Equals(locTok.TokenString, StringComparison.OrdinalIgnoreCase)) Then
                    currentModule.Variables.Add(New ParamInfo With {
                    .Name = tok.TokenString,
                    .DataType = tok.DataType,
                    .DeclaredInLine = tok.LineNumber
                })
                End If

            Case TokenTypeEnum.constant_decl
                If Not currentModule.Constants.Any(Function(c) c.Name.Equals(locTok.TokenString, StringComparison.OrdinalIgnoreCase)) Then
                    currentModule.Constants.Add(New ParamInfo With {
                    .Name = tok.TokenString,
                    .DataType = tok.DataType,
                    .DeclaredInLine = tok.LineNumber
                })
                End If

            Case TokenTypeEnum.declare_function, TokenTypeEnum.declare_sub
                If Not currentModule.Declares.ContainsKey(tok.TokenString) Then
                    Dim declInfo As New MethodInfo With {
                    .Name = tok.TokenString,
                    .ParentModule = currentModule.Name,
                    .MethodScope = "Declare"
                }
                    currentModule.Declares(tok.TokenString) = declInfo
                End If
        End Select
    End Sub

    Private Function BuildQualifiedWithName(line As MethodLine, tok As Token) As String
        If tok Is Nothing Then Return ""
        Dim qualifier As New List(Of String)
        Dim visitedTokens As New HashSet(Of Token)

        ' 1) Urcă prin părinți (expresii directe)
        Dim current As Token = tok
        While current IsNot Nothing AndAlso Not visitedTokens.Contains(current)
            visitedTokens.Add(current)
            If Not String.IsNullOrWhiteSpace(current.TokenString) Then
                qualifier.Insert(0, current.TokenString.Trim("."c))
            End If
            current = current.Parent
        End While

        ' 2) Adaugă prefixele din With (recursiv, suportă With în With)
        If line IsNot Nothing AndAlso tok.IsWithMember Then
            Dim ctx As MethodLine = line
            Dim visitedLines As New HashSet(Of Integer)

            While ctx IsNot Nothing AndAlso ctx.WithBlockContext IsNot Nothing
                Dim parentLine = ctx.WithBlockContext
                If visitedLines.Contains(parentLine.LineNumber) Then Exit While
                visitedLines.Add(parentLine.LineNumber)

                ' tokenul de bază după "With"
                Dim baseTok As Token =
                If(parentLine.Tokens.
                    SkipWhile(Function(t) Not t.TokenString.Equals("With", StringComparison.OrdinalIgnoreCase)).
                    Skip(1).
                    FirstOrDefault(Function(t) Not String.IsNullOrWhiteSpace(t.TokenString)), parentLine.Tokens.
                    FirstOrDefault(Function(tt) Not String.IsNullOrWhiteSpace(tt.TokenString) AndAlso
                                              Not tt.TokenString.Equals("With", StringComparison.OrdinalIgnoreCase)))

                If baseTok IsNot Nothing Then
                    ' 🔁 ia calificarea COMPLETĂ a bazei (poate fi și ea tot dintr-un With)
                    Dim baseQual = BuildQualifiedWithName(parentLine, baseTok)
                    If Not String.IsNullOrEmpty(baseQual) Then
                        qualifier.Insert(0, baseQual)
                    Else
                        qualifier.Insert(0, baseTok.TokenString.Trim("."c))
                    End If
                End If

                ctx = parentLine.WithBlockContext
            End While
        End If

        ' 3) Concatenează
        If qualifier.Count = 0 Then
            Return tok.TokenString.Trim("."c)
        Else
            Return String.Join(".", qualifier)
        End If
    End Function

    Friend Sub LogInfoLocal(ByVal msg As String, level As Integer, Optional force As Boolean = False)
        If pCurrentMethod.Value <> "" Then
            msg &= $"({pCurrentType.Value}) [{pCurrentModule.Value}].[{pCurrentMethod.Value}] : {msg}"
        Else
            msg &= $"({pCurrentType.Value}) [{pCurrentModule.Value}] : {msg}"
        End If

        LogInfo("[TOKENIZER]" & msg, level, force)
    End Sub

    ''' <summary>
    ''' Parsează o linie de tip etichetă (Label:) și o înregistrează în metoda curentă.
    ''' </summary>
    Private Function GetTokensFromLabelLine(currentMethod As MethodInfo, lineObj As MethodLine) As Token
        Try

            Dim labelName = lineObj.Content.TrimEnd(":"c)

            Dim labelTok As New Token With {
                .LineNumber = lineObj.LineNumber,
                .LocalLineNumber = lineObj.LocalLineNumber,
                .TokenString = labelName,
                .TokenType = TokenTypeEnum.label_decl,
                .IsResolved = True,
                .ResolvedScope = "label_decl",
                .DataType = "Label",
                .Context = currentMethod?.Name,
                .SourceLine = lineObj
            }

            Return labelTok

        Catch ex As Exception
            LogError($"ParseLabelLine({currentMethod?.Name})", ex)
            Return Nothing
        End Try
    End Function
End Module

Public Module TokenizerResolves
    ''' <summary>
    ''' Analizează toate metodele unui modul Form/Report și,
    ''' pentru fiecare metodă declarativă ce corespunde unui eveniment din EventSets,
    ''' marchează metoda ca handler de eveniment și generează un token asociat.
    ''' </summary>
    Friend Sub ResolveAccessEventHandlers(ByRef currentModule As FormReportContainer)
        Try
            If currentModule Is Nothing OrElse currentModule.EventSets Is Nothing OrElse currentModule.EventSets.Count = 0 Then Exit Sub
            If currentModule.Methods Is Nothing OrElse currentModule.Methods.Count = 0 Then Exit Sub

            For Each m In currentModule.Methods.Values
                Dim declLine As MethodLine = m.MethodLines?.FirstOrDefault(Function(l) l.StartsMethodBlock)
                If declLine Is Nothing Then Continue For

                Dim methodName = m.Name
                If String.IsNullOrEmpty(methodName) Then Continue For

                Dim lastUnderscore As Integer = methodName.LastIndexOf("_"c)
                If lastUnderscore <= 0 Then Continue For

                Dim ctrlName As String = methodName.Substring(0, lastUnderscore)
                Dim evtName As String = methodName.Substring(lastUnderscore + 1)

                Dim matchKey = currentModule.EventSets.Keys.FirstOrDefault(Function(k) k.Equals(methodName, StringComparison.OrdinalIgnoreCase))
                If String.IsNullOrEmpty(matchKey) Then Continue For

                Dim match = currentModule.EventSets(matchKey)
                If match Is Nothing Then Continue For

                ' === referință la obiectul real ===
                Dim resolvedObj As Object = Nothing
                Select Case match.RefType
                    Case "Control"
                        resolvedObj = currentModule.Controls.FirstOrDefault(Function(c) c.Name.Equals(match.HandlerString, StringComparison.OrdinalIgnoreCase))
                    Case "Section"
                        resolvedObj = currentModule.Sections.FirstOrDefault(Function(s) s.Name.Equals(match.HandlerString, StringComparison.OrdinalIgnoreCase))
                    Case "Form", "Report"
                        resolvedObj = currentModule
                End Select

                ' === marchează metoda ===
                m.IsAccessEventHandler = True
                m.HandlerObject = match.HandlerString
                m.MethodScope = "access_event"

                ' === creează tokenul evenimentului ===
                Dim evtTok As New Token With {
                    .TokenString = methodName,
                    .TokenType = TokenTypeEnum.access_event,
                    .IsResolved = True,
                    .ResolvedScope = $"{match.HandlerString}_event",
                    .ResolvedRef = resolvedObj,
                    .DataType = $"{match.RefType}.{match.EventName}",
                    .LineNumber = declLine.LineNumber,
                    .LocalLineNumber = declLine.LocalLineNumber,
                    .SourceLine = declLine,
                    .Context = m.Name
                }

                declLine.Tokens = New TokenList From {evtTok}
                MapTokensToColumns(declLine.Tokens, declLine.OriginalContent, declLine.Content)

                If m.Tokens Is Nothing Then m.Tokens = New TokenList()
                m.Tokens.Add(declLine.Tokens.FirstOrDefault)

                ' === marchează linia declarativă ===
                declLine.IsTokenized = True

                LogInfoLocal($"Access event matched: {methodName} → {match.RefType}.{match.EventName}", 2)
            Next

        Catch ex As Exception
            LogError("ResolveAccessEventHandlers", ex)
        End Try
    End Sub

    Friend Sub ResolveAccessEventHandler(ByRef tok As Token, ByRef currentMethod As MethodInfo, ByRef currentModule As ModuleContainer)
        Try
            If tok Is Nothing OrElse currentMethod Is Nothing OrElse currentModule Is Nothing Then Exit Sub
            If Not (TypeOf currentModule Is FormReportContainer) Then Exit Sub

            Dim frm = DirectCast(currentModule, FormReportContainer)
            Dim methodName = currentMethod.Name
            If String.IsNullOrEmpty(methodName) Then Exit Sub

            Dim parts = methodName.Split("_"c)
            'If parts.Length <> 2 Then Exit Sub

            Dim evtName As String = parts.Last()
            Dim ctrlName As String = String.Join("_", parts.Take(parts.Length - 1))

            ' === 1️⃣ Form_Load / Report_Open etc. ===
            If ctrlName.Equals("Form", StringComparison.OrdinalIgnoreCase) OrElse ctrlName.Equals("Report", StringComparison.OrdinalIgnoreCase) Then
                If modAccessPropertyCatalog.TryIsKnownFormEvent(evtName) Then
                    currentMethod.IsAccessEventHandler = True
                    currentMethod.HandlerObject = ctrlName
                    currentMethod.HandlerObjectType = currentModule.Type
                    currentMethod.HandlerEvent = evtName
                    currentMethod.MethodScope = "access_event"

                    ' --- Token ---
                    tok.TokenType = TokenTypeEnum.access_event
                    tok.ResolvedScope = $"{ctrlName}_event"
                    tok.DataType = $"{currentModule.Type}.{evtName}"
                    tok.IsResolved = True
                    tok.ResolvedRef = currentMethod

                    ' --- SymbolEntry ---
                    Dim sym As New SymbolEntry With {
                        .Name = methodName,
                        .TypeName = $"{currentModule.Type}.{evtName}",
                        .DeclType = TokenTypeEnum.access_event,
                        .Scope = "Private",
                        .SourceObject = currentMethod,
                        .Category = "AccessEvent"
                    }

                    If Not currentModule.ModuleSymbols.ContainsKey(sym.Name) Then currentModule.ModuleSymbols(sym.Name) = sym

                    LogInfoLocal($"Resolved Access event handler method: '{methodName}' for {currentModule.Type}", 3)
                    Exit Sub
                End If
            End If

            ' === 2️⃣ ControlName_EventName ===
            Dim ctrl = frm.Controls?.FirstOrDefault(Function(c) c.Name.Equals(ctrlName, StringComparison.OrdinalIgnoreCase))
            If ctrl Is Nothing Then Exit Sub

            If modAccessPropertyCatalog.TryIsKnownControlEvent(ctrl.Type, evtName) Then
                currentMethod.IsAccessEventHandler = True
                currentMethod.HandlerObject = ctrlName
                currentMethod.HandlerObjectType = ctrl.Type
                currentMethod.HandlerEvent = evtName
                currentMethod.MethodScope = "access_event"

                ' --- Token ---
                tok.TokenType = TokenTypeEnum.access_event
                tok.ResolvedScope = $"{ctrl.Type}_event"
                tok.DataType = $"{ctrl.Type}.{evtName}"
                tok.IsResolved = True
                tok.ResolvedRef = ctrl

                ' --- SymbolEntry ---
                Dim sym As New SymbolEntry With {
                    .Name = methodName,
                    .TypeName = $"{ctrl.Type}.{evtName}",
                    .DeclType = TokenTypeEnum.access_event,
                    .Scope = "Private",
                    .SourceObject = currentMethod,
                    .Category = "AccessEvent"
                }

                If Not currentModule.ModuleSymbols.ContainsKey(sym.Name) Then currentModule.ModuleSymbols(sym.Name) = sym
                LogInfoLocal($"Resolved Access event handler method: '{methodName}' for control '{ctrl.Name}' of type '{ctrl.Type}'", 3)
            End If

        Catch ex As Exception
            LogError("ResolveAccessEventHandler", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Rezolvă un token bazat pe contextul său (With sau chainParent).
    ''' </summary>
    Friend Sub ResolveMemberFromContext(ByRef tok As Token, contextObj As Object, currentModule As ModuleContainer)
        If tok Is Nothing OrElse contextObj Is Nothing Then Exit Sub

        Dim ctxTok As Token = TryCast(contextObj, Token)
        Dim frm As FormReportContainer = TryCast(currentModule, FormReportContainer)

        ' === 1️⃣ context "Me" (Form/Report curent) ===
        If ctxTok IsNot Nothing AndAlso ctxTok.TokenString.Equals("Me", StringComparison.OrdinalIgnoreCase) Then
            If frm IsNot Nothing Then
                ' a) Control pe form/report
                Dim symCtrl = TryFindMember(frm, tok.TokenString)
                If symCtrl IsNot Nothing Then
                    tok.IsResolved = True
                    tok.ResolvedRef = symCtrl
                    tok.DataType = symCtrl.TypeName
                    tok.TokenType = symCtrl.DeclType
                    tok.ResolvedScope = "chain_me_control"
                    LogInfoLocal($"Resolved token '{tok.TokenString}' as {symCtrl.TypeName} on form/report '{frm.Name}'", 3)
                    Exit Sub
                End If

                ' b) Proprietate de form
                If TryIsKnownFormProperty(tok.TokenString) Then
                    tok.IsResolved = True
                    tok.ResolvedRef = New SymbolEntry With {
                        .Name = tok.TokenString,
                        .TypeName = "Variant",
                        .DeclType = TokenTypeEnum.access_property,
                        .SourceObject = frm
                    }
                    tok.DataType = "Variant"
                    tok.TokenType = TokenTypeEnum.access_property
                    tok.ResolvedScope = "chain_me_property"
                    LogInfoLocal($"Resolved token '{tok.TokenString}' as form/report property on '{frm.Name}'", 3)
                    Exit Sub
                End If
            End If
        End If

        ' === 2️⃣ context "Control" (Access control) ===
        If ctxTok IsNot Nothing AndAlso ctxTok.TokenType = TokenTypeEnum.access_control Then
            ' === Obiectul din ResolvedRef poate fi direct FormReportControl sau un SymbolEntry care îl conține ===
            Dim ctrl As FormReportControl = Nothing

            If TypeOf ctxTok.ResolvedRef Is FormReportControl Then
                ctrl = DirectCast(ctxTok.ResolvedRef, FormReportControl)
            ElseIf TypeOf ctxTok.ResolvedRef Is SymbolEntry Then
                Dim se = DirectCast(ctxTok.ResolvedRef, SymbolEntry)
                ctrl = TryCast(se.SourceObject, FormReportControl)
            End If

            If ctrl IsNot Nothing Then
                Dim symProp = TryFindMember(ctrl, tok.TokenString)
                If symProp IsNot Nothing Then
                    tok.IsResolved = True
                    tok.ResolvedRef = symProp
                    tok.DataType = symProp.TypeName
                    tok.TokenType = symProp.DeclType
                    tok.ResolvedScope = "chain_control_property"
                    LogInfoLocal($"Resolved token '{tok.TokenString}' as property of control '{ctrl.Name}'", 3)
                    Exit Sub
                End If

                ' fallback: proprietate cunoscută de control
                If modAccessPropertyCatalog.TryIsKnownControlProperty(ctrl.Type, tok.TokenString) Then
                    tok.IsResolved = True
                    tok.ResolvedRef = New SymbolEntry With {
                        .Name = tok.TokenString,
                        .TypeName = "Variant",
                        .DeclType = TokenTypeEnum.access_property,
                        .SourceObject = ctrl
                    }
                    tok.DataType = "Variant"
                    tok.TokenType = TokenTypeEnum.access_property
                    tok.ResolvedScope = "chain_control_knownprop"
                    LogInfoLocal($"Resolved token '{tok.TokenString}' as known property of control '{ctrl.Name}'", 3)
                    Exit Sub
                End If
            End If
        End If

        ' === 3️⃣ context builtin (Err, DoCmd, etc.) ===
        If ctxTok IsNot Nothing AndAlso ctxTok.IsBuiltIn Then
            Dim symBuilt = TryFindMember(currentModule, tok.TokenString)
            If symBuilt IsNot Nothing Then
                tok.IsResolved = True
                tok.ResolvedRef = symBuilt
                tok.DataType = symBuilt.TypeName
                tok.TokenType = symBuilt.DeclType
                tok.ResolvedScope = "chain_builtin"
                Exit Sub
            End If

            tok.TokenType = TokenTypeEnum.builtin_property
            tok.IsResolved = True
            tok.ResolvedScope = "chain_builtin_fallback"
            LogInfoLocal($"Resolved token '{tok.TokenString}' as built-in fallback", 3)
            Exit Sub
        End If
    End Sub

    ''' <summary>
    ''' Rezolvă tokenii care depind de alți tokeni: lanțuri (a.b.c), blocuri WITH,
    ''' membri Enum/Type, proprietăți sau funcții definite în modul.
    ''' </summary>
    Friend Sub ResolveTokenRecursive(ByRef tok As Token, currentMethod As MethodInfo, currentModule As ModuleContainer)
        Try
            Dim locTok As Token = tok
            If tok Is Nothing OrElse tok.TokenType <> TokenTypeEnum.unresolved OrElse tok.IsBuiltIn Then Exit Sub

            ' === 1️⃣ Proprietăți de modul ===
            If currentModule?.Properties?.ContainsKey(tok.TokenString) Then
                Dim p = currentModule.Properties(tok.TokenString)
                tok.TokenType = TokenTypeEnum.property_get
                tok.DataType = p.ReturnType
                tok.ResolvedScope = "property"
                tok.IsResolved = True
                tok.ResolvedRef = p
                LogInfoLocal($"Resolved module property: '{tok.TokenString}'", 3)
                Exit Sub
            End If

            ' === 2️⃣ Funcții definite în modul ===
            If currentModule?.Functions?.ContainsKey(tok.TokenString) Then
                Dim f = currentModule.Functions(tok.TokenString)
                tok.TokenType = TokenTypeEnum.Function
                tok.DataType = f.ReturnType
                tok.ResolvedScope = "function"
                tok.IsResolved = True
                tok.ResolvedRef = f
                LogInfoLocal($"Resolved module function: '{tok.TokenString}'", 3)
                Exit Sub
            End If

            ' === 3️⃣ Enum / Type și membrii lor ===
            Dim et = If(currentModule?.Enums?.FirstOrDefault(Function(e) e.Name.Equals(locTok.TokenString, StringComparison.OrdinalIgnoreCase)), currentModule?.Types?.FirstOrDefault(Function(t) t.Name.Equals(locTok.TokenString, StringComparison.OrdinalIgnoreCase)))

            If et IsNot Nothing Then
                tok.TokenType = If(et.Type = "Enum", TokenTypeEnum.enum_type, TokenTypeEnum.user_type)
                tok.DataType = et.Name
                tok.ResolvedScope = "module_type"
                tok.IsResolved = True
                tok.ResolvedRef = et
                LogInfoLocal($"Resolved module type: '{tok.TokenString}'", 3)
                Exit Sub
            End If

            ' === 4️⃣ Membru de Enum/Type ===
            If tok.Parent IsNot Nothing AndAlso tok.Parent.IsResolved AndAlso TypeOf tok.Parent.ResolvedRef Is EnumOrType Then
                Dim pet = DirectCast(tok.Parent.ResolvedRef, EnumOrType)
                Dim member = pet.Members.FirstOrDefault(Function(m) m.Name.Equals(locTok.TokenString, StringComparison.OrdinalIgnoreCase))
                If member IsNot Nothing Then
                    tok.TokenType = If(pet.Type = "Enum", TokenTypeEnum.enum_member, TokenTypeEnum.type_field)
                    tok.DataType = If(pet.Type = "Enum", "Long", pet.Name)
                    tok.ResolvedScope = If(pet.Type = "Enum", "enum_member", "type_field")
                    tok.IsResolved = True
                    tok.ResolvedRef = member
                    tok.QualifiedName = $"{pet.Name}.{member.Name}"
                    LogInfoLocal($"Resolved {pet.Type} member: '{tok.TokenString}' of '{pet.Name}'", 3)
                    Exit Sub
                End If
            End If

            ' === 5️⃣ With block sau lanțuri Parent.Member ===
            Dim rootTok As Token = If(tok.IsWithMember, GetRootToken(tok), tok.Parent)
            If rootTok IsNot Nothing AndAlso rootTok IsNot tok Then
                If Not rootTok.IsResolved Then
                    ResolveTokenRecursive(rootTok, currentMethod, currentModule)
                End If

                If rootTok.IsResolved AndAlso rootTok.ResolvedRef IsNot Nothing Then
                    Dim parentObj = If(TypeOf rootTok.ResolvedRef Is SymbolEntry,
                           DirectCast(rootTok.ResolvedRef, SymbolEntry).SourceObject,
                           rootTok.ResolvedRef)

                    Dim child = TryFindMember(parentObj, tok.TokenString.TrimStart("."c))
                    If child IsNot Nothing Then
                        Select Case child.DeclType
                            Case TokenTypeEnum.property_get_decl, TokenTypeEnum.property_let_decl, TokenTypeEnum.property_set_decl
                                tok.TokenType = TokenTypeEnum.property_get
                            Case TokenTypeEnum.function_decl
                                tok.TokenType = TokenTypeEnum.Function
                            Case TokenTypeEnum.access_property
                                tok.TokenType = TokenTypeEnum.access_property
                            Case TokenTypeEnum.access_control
                                tok.TokenType = TokenTypeEnum.access_control
                            Case Else
                                tok.TokenType = TokenTypeEnum.variable_member
                        End Select

                        tok.DataType = child.TypeName
                        tok.ResolvedScope = If(tok.IsWithMember, "with_member", "member_chain")
                        tok.IsResolved = True
                        tok.ResolvedRef = child
                        LogInfoLocal($"Resolved member: '{tok.TokenString}' of '{rootTok.TokenString}'", 3)
                        Exit Sub
                    End If
                End If
            End If

            ' === 6️⃣ Dacă nu s-a putut rezolva ===
            tok.IsResolved = False
            tok.ResolvedScope = "unresolved"
            LogInfoLocal($"Token remains unresolved: '{tok.TokenString}'", 3)

        Catch ex As Exception
            LogError("ResolveTokenRecursive", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Determină dacă tokenul se referă la o variabilă locală, parametru,
    ''' câmp de clasă, variabilă de modul, constantă etc.
    ''' Nu face rezoluție recursivă sau lanțuri.
    ''' </summary>
    Friend Sub ResolveTokenScope(ByRef tok As Token, currentMethod As MethodInfo, currentModule As ModuleContainer)
        Try
            Dim locTok As Token = tok
            If tok Is Nothing OrElse tok.TokenType <> TokenTypeEnum.unresolved Then Exit Sub

            Dim name = tok.TokenString
            tok.IsResolved = False
            tok.ResolvedRef = Nothing
            tok.ResolvedScope = ""
            tok.DataType = ""

            ' === 1️⃣ Built-in ===
            If tok.IsBuiltIn Then Exit Sub

            Try
                ' === 2️⃣ Token = "Me" ===
                If name.Equals("Me", StringComparison.OrdinalIgnoreCase) Then
                    tok.IsResolved = True
                    tok.TokenType = TokenTypeEnum.self
                    tok.ResolvedRef = currentModule
                    tok.DataType = currentModule.Name

                    Select Case True
                        Case TypeOf currentModule Is FormReportContainer
                            tok.ResolvedScope = "form_instance"
                        Case currentModule.Type.Equals("ClassModule", StringComparison.OrdinalIgnoreCase)
                            tok.ResolvedScope = "class_instance"
                        Case Else
                            tok.ResolvedScope = "module_self"
                    End Select
                    LogInfoLocal($"Token 're' solved as 'Me' reference", 3)
                    Exit Sub
                End If

            Catch ex As Exception
                LogError("ResolveTokenScope.Me", ex)
            End Try

            ' === 3️⃣ Variabile / Parametri locali ===
            If currentMethod IsNot Nothing Then
                Try
                    Dim localVar = currentMethod.Variables?.FirstOrDefault(Function(v) v.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                    If localVar IsNot Nothing Then
                        tok.IsResolved = True
                        tok.TokenType = TokenTypeEnum.variable
                        tok.ResolvedScope = "local_var"
                        tok.DataType = localVar.Type
                        tok.ResolvedRef = localVar
                        LogInfoLocal($"Token '{tok.TokenString}' resolved as local variable", 3)
                        Exit Sub
                    End If

                Catch ex As Exception
                    LogError("ResolveTokenScope.LocalVar", ex)
                End Try

                Try
                    Dim param = currentMethod.Parameters?.FirstOrDefault(Function(p) p.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                    If param IsNot Nothing Then
                        tok.IsResolved = True
                        tok.TokenType = TokenTypeEnum.parameter
                        tok.ResolvedScope = "param"
                        tok.DataType = param.Type
                        tok.ResolvedRef = param
                        LogInfoLocal($"Token '{tok.TokenString}' resolved as parameter", 3)
                        Exit Sub
                    End If

                Catch ex As Exception
                    LogError("ResolveTokenScope.Parameter", ex)
                End Try
            End If

            Try
                ' === 4️⃣ Variabile / Constante de modul ===
                Dim vMod = currentModule?.Variables?.FirstOrDefault(Function(v) v.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                If vMod IsNot Nothing Then
                    tok.IsResolved = True
                    tok.TokenType = TokenTypeEnum.variable
                    tok.ResolvedScope = "module_var"
                    tok.DataType = vMod.Type
                    tok.ResolvedRef = vMod
                    LogInfoLocal($"Token '{tok.TokenString}' resolved as module variable", 3)
                    Exit Sub
                End If

            Catch ex As Exception
                LogError("ResolveTokenScope.ModuleVar", ex)
            End Try

            Try
                Dim cMod = currentModule?.Constants?.FirstOrDefault(Function(c) c.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                If cMod IsNot Nothing Then
                    tok.IsResolved = True
                    tok.TokenType = TokenTypeEnum.constant
                    tok.ResolvedScope = "module_const"
                    tok.DataType = cMod.Type
                    tok.ResolvedRef = cMod
                    LogInfoLocal($"Token '{tok.TokenString}' resolved as module constant", 3)
                    Exit Sub
                End If

            Catch ex As Exception
                LogError("ResolveTokenScope.ModuleConst", ex)
            End Try

            Try
                ' === 5️⃣ Simboluri declarate la nivel de modul ===
                If currentModule IsNot Nothing AndAlso currentModule.ModuleSymbols IsNot Nothing Then
                    Dim sym As SymbolEntry = Nothing
                    If currentModule.ModuleSymbols.TryGetValue(name, sym) Then
                        tok.IsResolved = True
                        tok.TokenType = sym.DeclType
                        tok.ResolvedScope = "module_symbol"
                        tok.ResolvedRef = sym.SourceObject
                        tok.DataType = sym.TypeName
                        LogInfoLocal($"Token '{tok.TokenString}' resolved as module symbol", 3)
                        Exit Sub
                    End If
                End If

            Catch ex As Exception
                LogError("ResolveTokenScope.ModuleSymbol", ex)
            End Try


            Try
                ' === 6️⃣ Enum / Type la nivel de modul (ca tip, nu membru) ===
                Dim et = If(currentModule?.Enums?.FirstOrDefault(Function(e) e.Name.Equals(name, StringComparison.OrdinalIgnoreCase)), currentModule?.Types?.FirstOrDefault(Function(t) t.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))

                If et IsNot Nothing Then
                    tok.IsResolved = True
                    tok.TokenType = If(et.Type = "Enum", TokenTypeEnum.enum_type, TokenTypeEnum.user_type)
                    tok.ResolvedScope = "module_type"
                    tok.DataType = et.Name
                    tok.ResolvedRef = et
                    LogInfoLocal($"Token '{tok.TokenString}' resolved as module type", 3)
                    Exit Sub
                End If

            Catch ex As Exception
                LogError("ResolveTokenScope.ModuleType", ex)
            End Try

            ' === Caută un token declarat anterior în aceeași metodă ===
            If currentMethod IsNot Nothing Then
                Try
                    If currentMethod.CodeLabels IsNot Nothing Then
                        Dim lbl = currentMethod.CodeLabels?.FirstOrDefault(Function(l) l.TokenString.Equals(locTok.TokenString, StringComparison.OrdinalIgnoreCase))

                        If lbl IsNot Nothing Then
                            tok.TokenType = TokenTypeEnum.label_ref
                            tok.ResolvedScope = "label_ref"
                            tok.IsResolved = True
                            tok.DataType = "CodeLabel"
                            tok.ResolvedRef = lbl
                            LogInfoLocal($"Token '{tok}' resolved as code label reference", 3)
                            Exit Sub
                        End If
                    End If

                Catch ex As Exception
                    LogError("ResolveTokenScope.LabelRef", ex)
                End Try

                Try
                    Dim declTok = currentMethod.Tokens?.FirstOrDefault(Function(t) _
                                                                (t.TokenType = TokenTypeEnum.variable_decl OrElse
                                                                 t.TokenType = TokenTypeEnum.parameter_decl OrElse
                                                                 t.TokenType = TokenTypeEnum.constant_decl) AndAlso
                                                                 t.TokenString.Equals(name, StringComparison.OrdinalIgnoreCase))

                    If declTok IsNot Nothing Then
                        tok.TokenType = TokenTypeEnum.variable
                        tok.DataType = declTok.DataType
                        tok.ResolvedRef = declTok
                        tok.IsResolved = True
                        tok.ResolvedScope = "local_decl_ref"
                        LogInfoLocal($"Token '{tok}' resolved as local declaration reference", 3)
                        Exit Sub
                    End If

                Catch ex As Exception
                    LogError("ResolveTokenScope.LocalDeclRef", ex)
                End Try
            End If

            ' === 7️⃣ Dacă n-a fost găsit, rămâne nerezolvat ===
            tok.ResolvedScope = "unresolved"

        Catch ex As Exception
            LogError("ResolveTokenScope", ex)
        End Try
    End Sub

    Private Function TryFindMember(obj As Object, memberName As String) As SymbolEntry
        If obj Is Nothing OrElse String.IsNullOrEmpty(memberName) Then Return Nothing

        ' === 1️⃣ ModuleContainer ===
        If TypeOf obj Is ModuleContainer Then
            Dim modObj = DirectCast(obj, ModuleContainer)
            Dim sym As SymbolEntry = Nothing
            If modObj.ModuleSymbols IsNot Nothing AndAlso modObj.ModuleSymbols.TryGetValue(memberName, sym) Then
                Return sym
            End If
        End If

        ' === 2️⃣ FormReportContainer (include Controls + form properties) ===
        If TypeOf obj Is FormReportContainer Then
            Dim frm = DirectCast(obj, FormReportContainer)

            ' (a) Proprietăți de form
            If TryIsKnownFormProperty(memberName) Then
                Return New SymbolEntry With {
                .Name = memberName,
                .TypeName = "Variant",
                .DeclType = TokenTypeEnum.access_property,
                .SourceObject = frm
            }
            End If

            ' (b) Membri normali declarați în modulul formului
            Dim sym As SymbolEntry = Nothing
            If frm.ModuleSymbols IsNot Nothing AndAlso frm.ModuleSymbols.TryGetValue(memberName, sym) Then
                Return sym
            End If

            ' (c) Controale — deja încărcate în ModuleSymbols, dar fallback direct
            Dim ctrl = frm.Controls?.FirstOrDefault(Function(c) c.Name.Equals(memberName, StringComparison.OrdinalIgnoreCase))
            If ctrl IsNot Nothing Then
                Return New SymbolEntry With {
                .Name = ctrl.Name,
                .TypeName = ctrl.Type,
                .DeclType = TokenTypeEnum.access_control,
                .SourceObject = ctrl
            }
            End If
        End If

        ' === 3️⃣ FormReportControl (folosește catalogul de proprietăți implicite Access) ===
        If TypeOf obj Is FormReportControl Then
            Dim ctrl = DirectCast(obj, FormReportControl)

            If modAccessPropertyCatalog.TryIsKnownControlProperty(ctrl.Type, memberName) Then
                Return New SymbolEntry With {
                .Name = memberName,
                .TypeName = "Variant",
                .DeclType = TokenTypeEnum.access_property,
                .SourceObject = ctrl
            }
            End If
        End If

        ' === 4️⃣ Alte obiecte (PropertyInfo, FunctionInfo etc.) ===
        Try
            Dim t = obj.GetType()
            Dim prop = t.GetProperty(memberName)
            If prop IsNot Nothing Then
                Return New SymbolEntry With {
                .Name = memberName,
                .TypeName = prop.PropertyType.Name,
                .DeclType = TokenTypeEnum.external_prefix,
                .SourceObject = prop
            }
            End If
        Catch
        End Try

        Return Nothing
    End Function

    ''' <summary>
    ''' Returnează tokenul-rădăcină (fără Parent) pentru expresii normale,
    ''' sau tokenul principal din blocul With corespunzător (ex: "Me" / "db").
    ''' </summary>
    Private Function GetRootToken(tok As Token) As Token
        If tok Is Nothing Then Return Nothing

        ' 🔹 1️⃣ Urcă lanțul Parent (pentru expresii tip a.b.c)
        Dim current As Token = tok
        While current.Parent IsNot Nothing
            current = current.Parent
        End While

        ' 🔹 2️⃣ Dacă e în With-block, urcă la tokenul de bază al acelui With
        If tok.IsWithMember AndAlso tok.WithBlockContext IsNot Nothing Then
            Dim parentLine = tok.WithBlockContext

            Dim baseTok As Token =
            If(parentLine.Tokens.
                SkipWhile(Function(t) Not t.TokenString.Equals("With", StringComparison.OrdinalIgnoreCase)).
                Skip(1).
                FirstOrDefault(Function(t) Not String.IsNullOrWhiteSpace(t.TokenString)), parentLine.Tokens.
                FirstOrDefault(Function(tt) Not String.IsNullOrWhiteSpace(tt.TokenString) AndAlso
                                           Not tt.TokenString.Equals("With", StringComparison.OrdinalIgnoreCase)))

            If baseTok IsNot Nothing Then
                Return GetRootToken(baseTok) ' 🔁 urcă recursiv
            End If
        End If

        Return current
    End Function
End Module