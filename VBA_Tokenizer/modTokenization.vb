'PROJECT NAME: VBA_TOKENIZER
'FILE DESCRIPTION: Module pentru tokenizarea codului VBA
'PATH: VBA_TOKENIZER/modTokenization.vb

Imports System.Text.RegularExpressions
Imports VBA_CORE
Imports VBA_CORE.modAccessPropertyCatalog
Imports VBA_CORE.modTypes

Public Module Tokenizer
    Private pCurrentType As New Threading.ThreadLocal(Of String)(Function() "")
    Private pCurrentModule As New Threading.ThreadLocal(Of String)(Function() "")
    Private pCurrentMethod As New Threading.ThreadLocal(Of String)(Function() "")

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
        If modObj Is Nothing OrElse modObj.Methods.Count = 0 Then Exit Sub

        pCurrentModule.Value = modObj.Name

        Dim opts As New ParallelOptions With {.MaxDegreeOfParallelism = Environment.ProcessorCount}
        'If modObj.Name = "clsDatabase" Then Stop

        ' === Tokenize fiecare metodă ===
        Dim methodAction As Action(Of MethodInfo) =
        Sub(m)
            ' Stivă locală pentru blocuri WITH (poate fi extinsă)
            Dim withStack As New Stack(Of MethodLine)

            For Each ml In m.MethodLines
                ' === gestionează blocurile WITH ===
                If ml.StartsWithBlock Then
                    withStack.Push(ml)
                ElseIf ml.EndsWithBlock Then
                    If withStack.Count > 0 Then withStack.Pop()
                ElseIf ml.IsInWithBlock AndAlso withStack.Count > 0 Then
                    ml.WithBlockContext = withStack.Peek()
                End If

                ' === tokenizează linia ===
                Dim toks = GetTokensFromLine(m, modObj, ml)
                ml.Tokens = toks
                m.Tokens.AddRange(toks)
            Next
        End Sub

        If modGlobals.Global_UseParallel Then
            Parallel.ForEach(modObj.Methods.Values, opts, methodAction)
        Else
            For Each m In modObj.Methods.Values
                methodAction(m)
            Next
        End If

        ' === Tokenizează liniile la nivel de modul (în afara metodelor) ===
        For Each ml In modObj.Lines
            If Not ml.StartsMethodBlock AndAlso Not ml.EndsMethodBlock Then
                Dim toks = GetTokensFromLine(Nothing, modObj, ml)
                ml.Tokens = toks
            End If
        Next
    End Sub


    ' ==============================================================
    ' 🔹 Extrage tokeni dintr-o linie curățată
    ' ==============================================================
    Private Function GetTokensFromLine(currentMethod As MethodInfo, currentModule As ModuleContainer, line As MethodLine, Optional methodContext As String = "") As List(Of Token)
        Dim tokens As New List(Of Token)
        Dim lastToken As modTypes.Token = Nothing
        Dim prevOp As String = Nothing
        Dim asContextTarget As modTypes.Token = Nothing

        If currentMethod IsNot Nothing Then
            pCurrentMethod.Value = currentMethod.Name
        Else
            pCurrentMethod.Value = ""
        End If

        'If pCurrentMethod.Value = "Form_Load" Then Stop
        If line.Content.Contains("Public Enum UserRole") Then Stop

        Try
            ' === 1️⃣ Obține conținutul liniei ===
            Dim cleanLine As String = line.Content
            If String.IsNullOrWhiteSpace(cleanLine) Then Return tokens

            ' === 2️⃣ Detectează contextul WITH (dacă linia e în bloc WITH) ===
            Dim withContext As MethodLine = Nothing
            Dim withTokenContext As Token = Nothing

            If line.IsInWithBlock AndAlso line.WithBlockContext IsNot Nothing Then
                withContext = line.WithBlockContext
            End If

            ' === 3️⃣ Parcurge tokenii din linie ===
            For Each m As Match In RegexCache.RxTokenizer.Matches(cleanLine)
                Dim fullExpression = m.Groups(1).Value
                Dim op = m.Groups(2).Value

                If String.IsNullOrWhiteSpace(fullExpression) Then Continue For
                If Char.IsDigit(fullExpression(0)) Then Continue For

                ' === 4️⃣ Sparge expresia compusă (ex: pSettings.Add) ===
                Dim parts() As String = fullExpression.Split("."c)
                Dim startsWithDot As Boolean = fullExpression.StartsWith(".")

                If startsWithDot Then
                    parts = parts.Where(Function(p) Not String.IsNullOrEmpty(p)).ToArray()
                End If

                ' === 5️⃣ Creează tokeni pentru fiecare parte ===
                Dim chainParent As modTypes.Token = Nothing

                For i As Integer = 0 To parts.Length - 1
                    Dim partName = parts(i).Replace("()", "").Trim()
                    If partName = "" Then Continue For

                    Dim tok As New modTypes.Token With {
                        .LineNumber = line.LineNumber,
                        .LocalLineNumber = line.LocalLineNumber,
                        .MethodLineNumber = i + 1,
                        .TokenString = partName,
                        .NextSymbol = If(i < parts.Length - 1, ".", op),
                        .Context = methodContext,
                        .SourceLine = line,
                        .WithBlockDepth = line.WithBlockDepth,
                        .TokenType = TokenTypeEnum.Unresolved
                    }

                    'If tok.TokenString = "Err" Then Stop

                    ' === 6️⃣ Context “AS” activ ===
                    If asContextTarget IsNot Nothing Then
                        If tok.TokenString.StartsWith(".") Then
                            asContextTarget.DataType &= tok.TokenString
                        Else
                            asContextTarget.DataType = tok.TokenString
                            If line.IsDeclarationLine Then
                                asContextTarget.TokenType = TokenTypeEnum.Variable
                            ElseIf line.StartsMethodBlock Then
                                asContextTarget.TokenType = TokenTypeEnum.Parameter
                            ElseIf line.IsInTypeBlock Then
                                asContextTarget.TokenType = TokenTypeEnum.Type
                            ElseIf line.IsInEnumBlock Then
                                asContextTarget.TokenType = TokenTypeEnum.Enum
                            ElseIf line.IsDeclarationLine AndAlso currentMethod Is Nothing Then
                                asContextTarget.TokenType = TokenTypeEnum.Parameter
                            ElseIf line.IsEventLine Then
                                asContextTarget.TokenType = TokenTypeEnum.Parameter
                            Else
                                Stop
                            End If
                            asContextTarget = Nothing
                        End If

                        tok.TokenType = TokenTypeEnum.Unresolved
                    End If

                    Dim testBuiltIn As Boolean = Not tok.NextSymbol.ToUpper().Contains("AS")

                    'If testBuiltIn AndAlso currentMethod IsNot Nothing Then
                    'testBuiltIn = currentMethod.Tokens.Contains()
                    'End If

                    If Not tok.NextSymbol.ToUpper().Contains("AS") Then
                        ' === 7️⃣ Determină tipul tokenului (simplificat, completăm ulterior) ===
                        Dim category = modIgnoreLists.GetIgnoreCategory(partName)
                        Select Case category
                            Case "VBA Type" : tok.TokenType = "builtin_type"
                            Case "VBA Function" : tok.TokenType = "builtin_function"
                            Case "VBA Object" : tok.TokenType = "builtin_object"
                            Case "VBA Constant" : tok.TokenType = "builtin_constant"
                            Case "Access Object" : tok.TokenType = "access_object"
                            Case "Access Function" : tok.TokenType = "access_function"
                            Case "External Library" : tok.TokenType = "external_prefix"
                            Case "VBA Keyword" : tok.TokenType = "vba_keyword"
                            Case Else : tok.TokenType = TokenTypeEnum.Unresolved
                        End Select
                    End If

                    ' === 7.1️⃣ Detectează tokenii din linii de declarație (Dim / Const) ===
                    If line.IsDeclarationLine AndAlso tok.TokenType = TokenTypeEnum.Unresolved Then
                        ' Verificăm tipul de declarație
                        Dim isConstDecl = line.Content.TrimStart().StartsWith("Const ", StringComparison.OrdinalIgnoreCase)

                        ' Dacă tokenul urmează de obicei după Dim/Const și nu e keyword
                        If Not tok.TokenString.Equals("Dim", StringComparison.OrdinalIgnoreCase) AndAlso Not tok.TokenString.Equals("Const", StringComparison.OrdinalIgnoreCase) Then

                            ' Caz 1: dacă urmează "As" => tipul va fi completat mai jos
                            If tok.NextSymbol.Equals("As", StringComparison.OrdinalIgnoreCase) Then
                                tok.TokenType = If(isConstDecl, TokenTypeEnum.Constant, TokenTypeEnum.Variable)
                                tok.ResolvedScope = If(currentMethod Is Nothing, "module_decl", "local_decl")
                                tok.IsResolved = True
                                tok.DataType = ""          ' completăm la pasul cu AsContextTarget

                                ' Caz 2: dacă urmează virgulă sau nimic => implicit Variant
                            ElseIf tok.NextSymbol = "," OrElse tok.NextSymbol = "" Then
                                tok.TokenType = If(isConstDecl, TokenTypeEnum.Constant, TokenTypeEnum.Variable)
                                tok.ResolvedScope = If(currentMethod Is Nothing, "module_decl", "local_decl")
                                tok.IsResolved = True
                                tok.DataType = "Variant"
                            End If
                        End If
                    End If


                    ' === Ajustare specială pentru "Me" ===
                    If tok.TokenString.Equals("Me", StringComparison.OrdinalIgnoreCase) Then
                        ' în clase, forms și reports — "Me" este o instanță validă, nu built-in
                        If currentModule.Type.Equals("Class", StringComparison.OrdinalIgnoreCase) Then
                            tok.IsBuiltIn = False
                            tok.TokenType = currentModule.Name
                        Else
                            tok.IsBuiltIn = True
                        End If
                    Else
                        tok.IsBuiltIn = (tok.TokenType <> TokenTypeEnum.Unresolved)
                    End If

                    ' 🔹 stabilește relația internă în expresie (Parent → Child)
                    If chainParent IsNot Nothing Then tok.Parent = chainParent

                    ' 🔹 context “WITH” (doar dacă expresia începe cu punct)
                    If withContext IsNot Nothing AndAlso startsWithDot Then
                        tok.IsWithMember = True
                        tok.WithBlockContext = withContext
                        tok.ResolvedScope = "with_context"
                    End If

                    ' 🔹 enum or type
                    'If line.IsInEnumBlock Or line.IsInTypeBlock Then
                    '    Dim matchMember = currentModule.Types.FirstOrDefault(Function(t) t.Members.Any(Function(mm) mm.Name.Equals(tok.TokenString, StringComparison.OrdinalIgnoreCase)))

                    '    If matchMember IsNot Nothing Then
                    '        Dim match = matchMember.Members.First(Function(mm) mm.Name.Equals(tok.TokenString, StringComparison.OrdinalIgnoreCase))
                    '        tok.QualifiedName = match.QualifiedName
                    '        tok.ResolvedScope = If(line.IsInEnumBlock, "enum_member", "type_field")
                    '        tok.DataType = If(line.IsInEnumBlock, "Long", matchMember.Name)
                    '        tok.IsResolved = True
                    '    End If
                    'End If

                    ' 🔹 construiește QualifiedName DOAR dacă există Parent sau WithContext
                    If tok.IsWithMember Then
                        tok.QualifiedName = BuildQualifiedWithName(line, tok)
                        'tok.Parent = withContext.Tokens.FirstOrDefault()
                        tok.ParentType = withContext.Tokens.FirstOrDefault().TokenType
                        If withContext.Tokens.FirstOrDefault() IsNot Nothing AndAlso withContext.Tokens.FirstOrDefault().IsBuiltIn Then
                            tok.IsResolved = True
                            tok.TokenType = "builtin_property"
                            tokens.Add(tok)
                            chainParent = tok
                            Continue For
                        End If
                    End If

                    ' === 8️⃣ Doar dacă e un token „rezolvabil” (nu built-in / keyword)
                    If Not tok.IsBuiltIn Then
                        ' 🔹 marchează ca nerezolvat inițial
                        tok.IsResolved = False

                        ' 🔹 adaugă în liste
                        tokens.Add(tok)

                        ' 🔹 actualizează părintele intern pentru următoarea parte din expresie
                        chainParent = tok

                        ' 🔹 adaugă în contexte / rezolvă
                        AddTokenToMethodContext(tok, currentMethod)
                        AddTokenToModuleContext(tok, currentModule)
                        If Not tok.NextSymbol.ToUpper.Contains("AS") Then
                            ResolveTokenScope(tok, currentMethod, currentModule)
                        Else
                            asContextTarget = tok
                        End If
                    Else
                        If tok.TokenType <> "vba_keyword" And tok.TokenType <> "vba_constant" Then
                            tokens.Add(tok)
                            ' 🔹 actualizează părintele intern pentru următoarea parte din expresie
                            chainParent = tok
                        End If
                    End If

                    If line.StartsWithBlock Then
                        withTokenContext = tok
                    End If
                Next

                prevOp = op
            Next

            ' === 13️⃣ Salvează tokenii la nivel de linie ===
            line.Tokens = tokens

            ' === 14️⃣ Mapează pozițiile caracterelor în linia originală ===
            MapTokensToColumns(tokens, line.OriginalContent, cleanLine)

            ' =============================================================
            '   🔹 Context-aware resolution pentru funcții și proprietăți
            ' =============================================================
            For Each tok In tokens
                If tok.TokenType <> "unresolved" Or tok.IsBuiltIn Then Continue For
                ResolveTokenRecursive(tok, currentMethod, currentModule)
            Next

            ' === 🔹 Marcare explicită a tokenului din linia de declarație a unui EventHandler ===
            If currentMethod IsNot Nothing AndAlso
               currentMethod.IsAccessEventHandler AndAlso
               line Is currentMethod.MethodLines.First() Then

                Dim firstTok = tokens.FirstOrDefault(Function(t) t.TokenString.Equals(currentMethod.Name, StringComparison.OrdinalIgnoreCase))

                If firstTok IsNot Nothing Then
                    firstTok.TokenType = "event_handler"
                    firstTok.ResolvedScope = "access_event"
                    firstTok.DataType = "EventHandler"
                    firstTok.IsResolved = True

                    LogInfoLocal($" ↳ Token tagged as Access EventHandler: {currentMethod.HandlerObject}.{currentMethod.HandlerEvent} ({currentMethod.HandlerObjectType})", 2)
                End If
            End If

            Return tokens

        Catch ex As Exception
            LogError("GetTokensFromLine", ex)
            Return New List(Of Token)
        End Try
    End Function

    ''' <summary>
    ''' Încearcă să rezolve un token într-un anumit context (metodă + modul).
    ''' Poate fi apelată recursiv pentru lanțuri (ex: obj.Property.Method).
    ''' </summary>
    Private Sub ResolveTokenRecursive(tok As Token, currentMethod As MethodInfo, currentModule As ModuleContainer)
        If tok Is Nothing OrElse tok.TokenType <> "unresolved" OrElse tok.IsBuiltIn Then Exit Sub

        ' === 1️⃣ Variabile locale ===
        Dim vLoc = currentMethod?.Variables?.
        FirstOrDefault(Function(v) v.Name.Equals(tok.TokenString, StringComparison.OrdinalIgnoreCase))
        If vLoc IsNot Nothing Then
            tok.TokenType = "variable"
            tok.DataType = vLoc.Type
            tok.ResolvedScope = "local_var"
            tok.IsResolved = True
            tok.ResolvedRef = vLoc
            Exit Sub
        End If

        ' === 2️⃣ Parametrii locali ===
        Dim param = currentMethod?.Parameters?.
        FirstOrDefault(Function(p) p.Name.Equals(tok.TokenString, StringComparison.OrdinalIgnoreCase))
        If param IsNot Nothing Then
            tok.TokenType = "parameter"
            tok.DataType = param.Type
            tok.ResolvedScope = "param"
            tok.IsResolved = True
            tok.ResolvedRef = param
            Exit Sub
        End If

        ' === 3️⃣ Proprietăți definite în modul ===
        If currentModule?.Properties?.ContainsKey(tok.TokenString) Then
            Dim p = currentModule.Properties(tok.TokenString)
            tok.TokenType = "property"
            tok.DataType = p.ReturnType
            tok.ResolvedScope = "property"
            tok.IsResolved = True
            If p.PropertyType.Equals("Let", StringComparison.OrdinalIgnoreCase) Then tok.IsLeftSide = True
            Exit Sub
        End If

        ' === 4️⃣ Funcții definite în modul ===
        If currentModule?.Functions?.ContainsKey(tok.TokenString) Then
            Dim f = currentModule.Functions(tok.TokenString)
            tok.TokenType = "function"
            tok.DataType = f.ReturnType
            tok.ResolvedScope = "function"
            tok.IsResolved = True
            If Not String.IsNullOrEmpty(f.ReturnType) Then tok.IsReturnValue = True
            Exit Sub
        End If

        ' === 5️⃣ Variabile la nivel de modul ===
        Dim vMod = currentModule?.Variables?.
        FirstOrDefault(Function(v) v.Name.Equals(tok.TokenString, StringComparison.OrdinalIgnoreCase))
        If vMod IsNot Nothing Then
            tok.TokenType = "variable"
            tok.DataType = vMod.Type
            tok.ResolvedScope = "module_var"
            tok.IsResolved = True
            tok.ResolvedRef = vMod
            Exit Sub
        End If

        ' === 6️⃣ Enum sau Type (ca tip direct) ===
        Dim enumOrType = currentModule?.Enums?.
        FirstOrDefault(Function(e) e.Name.Equals(tok.TokenString, StringComparison.OrdinalIgnoreCase))
        If enumOrType Is Nothing Then
            enumOrType = currentModule?.Types?.
            FirstOrDefault(Function(t) t.Name.Equals(tok.TokenString, StringComparison.OrdinalIgnoreCase))
        End If

        If enumOrType IsNot Nothing Then
            tok.TokenType = enumOrType.Type.ToLowerInvariant()   ' "enum" / "type"
            tok.DataType = enumOrType.Name
            tok.ResolvedScope = "module_type"
            tok.IsResolved = True
            tok.ResolvedRef = enumOrType
            Exit Sub
        End If

        ' === 7️⃣ Membru de Enum/Type (ex: SomeEnum.Value / SomeType.Field) ===
        If tok.Parent IsNot Nothing Then
            Dim parentTok = tok.Parent
            ' Rezolvăm părintele dacă nu e rezolvat deja
            If Not parentTok.IsResolved Then
                ResolveTokenRecursive(parentTok, currentMethod, currentModule)
            End If

            ' Dacă părintele este un Enum / Type
            If parentTok.IsResolved AndAlso TypeOf parentTok.ResolvedRef Is EnumOrType Then
                Dim et = DirectCast(parentTok.ResolvedRef, EnumOrType)
                Dim member = et.Members.
                FirstOrDefault(Function(m) m.Name.Equals(tok.TokenString, StringComparison.OrdinalIgnoreCase))
                If member IsNot Nothing Then
                    tok.TokenType = If(et.Type = "Enum", "enum_member", "type_field")
                    tok.DataType = If(et.Type = "Enum", "Long", et.Name)
                    tok.ResolvedScope = et.Type.ToLowerInvariant() & "_member"
                    tok.IsResolved = True
                    tok.ResolvedRef = member
                    tok.QualifiedName = $"{et.Name}.{member.Name}"
                    Exit Sub
                End If
            End If
        End If

        ' === 8️⃣ Blocuri WITH sau lanțuri ===
        If tok.IsWithMember Then
            Dim rootTok As Token = GetRootToken(tok)
            If rootTok IsNot Nothing AndAlso rootTok IsNot tok Then
                ResolveTokenRecursive(rootTok, currentMethod, currentModule)

                If rootTok.IsResolved Then
                    tok.TokenType = "variable_member"
                    tok.DataType = rootTok.DataType
                    tok.ResolvedScope = "with_member"
                    tok.IsResolved = True
                    tok.ResolvedRef = rootTok.ResolvedRef
                    Exit Sub
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Adaugă tokenii care definesc variabile, parametri sau constante în contextul metodei curente.
    ''' </summary>
    Private Sub AddTokenToMethodContext(tok As Token, currentMethod As MethodInfo)
        If currentMethod Is Nothing Then Exit Sub

        Select Case tok.TokenType
            Case "variable_decl"
                ' Dacă variabila nu există deja, o adăugăm
                If Not currentMethod.Variables.Any(Function(v) v.Name.Equals(tok.TokenString, StringComparison.OrdinalIgnoreCase)) Then
                    currentMethod.Variables.Add(New ParamInfo With {
                    .Name = tok.TokenString,
                    .DataType = tok.DataType,
                    .DeclaredInLine = tok.LineNumber
                })
                End If

            Case "parameter_decl"
                If Not currentMethod.Parameters.Any(Function(p) p.Name.Equals(tok.TokenString, StringComparison.OrdinalIgnoreCase)) Then
                    currentMethod.Parameters.Add(New ParamInfo With {
                    .Name = tok.TokenString,
                    .DataType = tok.DataType,
                    .DeclaredInLine = tok.LineNumber
                })
                End If

            Case "constant_decl"
                If Not currentMethod.Constants.Any(Function(c) c.Name.Equals(tok.TokenString, StringComparison.OrdinalIgnoreCase)) Then
                    currentMethod.Constants.Add(New ParamInfo With {
                    .Name = tok.TokenString,
                    .DataType = tok.DataType,
                    .DeclaredInLine = tok.LineNumber
                })
                End If
        End Select
    End Sub

    ''' <summary>
    ''' Adaugă tokenii declarați la nivel de modul (variabile, constante, declarații API, metode publice).
    ''' </summary>
    Private Sub AddTokenToModuleContext(tok As Token, currentModule As ModuleContainer)
        If currentModule Is Nothing Then Exit Sub

        Select Case tok.TokenType
            Case "variable_decl"
                If Not currentModule.Variables.Any(Function(v) v.Name.Equals(tok.TokenString, StringComparison.OrdinalIgnoreCase)) Then
                    currentModule.Variables.Add(New ParamInfo With {
                    .Name = tok.TokenString,
                    .DataType = tok.DataType,
                    .DeclaredInLine = tok.LineNumber
                })
                End If

            Case "constant_decl"
                If Not currentModule.Constants.Any(Function(c) c.Name.Equals(tok.TokenString, StringComparison.OrdinalIgnoreCase)) Then
                    currentModule.Constants.Add(New ParamInfo With {
                    .Name = tok.TokenString,
                    .DataType = tok.DataType,
                    .DeclaredInLine = tok.LineNumber
                })
                End If

            Case "declare_function", "declare_sub"
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

    ''' <summary>
    ''' Determină dacă tokenul curent se referă la o variabilă locală, parametru, câmp de clasă sau rămâne nerezolvat.
    ''' </summary>
    Public Sub ResolveTokenScope(tok As Token, currentMethod As MethodInfo, currentContainer As ModuleContainer)
        Try
            If tok.TokenType <> "unresolved" Then Exit Sub

            ' helper local
            Dim unwrap As Func(Of Object, Object) =
                Function(o As Object)
                    If TypeOf o Is SymbolEntry Then
                        Dim se = DirectCast(o, SymbolEntry)
                        Return If(se.SourceObject, o)
                    End If
                    Return o
                End Function

            Dim name = tok.TokenString
            tok.IsResolved = False
            tok.ResolvedRef = Nothing
            tok.ResolvedScope = ""
            tok.DataType = ""

            ' === 1️⃣ Built-in ===
            If tok.IsBuiltIn Then Exit Sub

            ' === 2️⃣ Token = "Me" ===
            If name.Equals("Me", StringComparison.OrdinalIgnoreCase) Then
                tok.IsResolved = True
                tok.ResolvedRef = currentContainer
                tok.DataType = currentContainer.Name

                If TypeOf currentContainer Is FormReportContainer Then
                    tok.ResolvedScope = "form_instance"
                ElseIf currentContainer.Type.Equals("ClassModule", StringComparison.OrdinalIgnoreCase) Then
                    tok.ResolvedScope = "class_instance"
                Else
                    tok.ResolvedScope = "module_self"
                End If

                Exit Sub
            End If

            ' === 3️⃣ Local variables / parameters ===
            If currentMethod IsNot Nothing Then
                Dim localVar = currentMethod.Variables?.FirstOrDefault(Function(v) v.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                If localVar IsNot Nothing Then
                    tok.IsResolved = True
                    tok.ResolvedScope = "local_var"
                    tok.ResolvedRef = localVar
                    tok.DataType = localVar.Type
                    tok.TokenType = "local_var"
                    Exit Sub
                End If

                Dim param = currentMethod.Parameters?.FirstOrDefault(Function(p) p.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                If param IsNot Nothing Then
                    tok.IsResolved = True
                    tok.ResolvedScope = "param"
                    tok.ResolvedRef = param
                    tok.DataType = param.Type
                    Exit Sub
                End If
            End If

            ' === 4️⃣ With block context (recursive resolution) ===
            If tok.IsWithMember AndAlso tok.WithBlockContext IsNot Nothing Then
                Dim parentLine = tok.WithBlockContext
                Dim withTokens = parentLine.Tokens

                ' 4.a) Dacă tokenul precedent din același lanț e deja rezolvat,
                '      folosește-l ca PĂRINTE pentru membrul curent (chaining corect).
                If tok.Parent IsNot Nothing AndAlso tok.Parent.IsResolved AndAlso tok.Parent.ResolvedRef IsNot Nothing Then
                    Dim parentObj = unwrap(tok.Parent.ResolvedRef)
                    Dim memberName = tok.TokenString.TrimStart("."c)
                    Dim childSym = TryFindMember(parentObj, memberName)

                    If childSym IsNot Nothing Then
                        tok.IsResolved = True
                        tok.ResolvedScope = "with_member"
                        tok.ResolvedRef = childSym
                        tok.DataType = childSym.TypeName
                        tok.TokenType = If(childSym.DeclType Is Nothing, childSym.Category, childSym.DeclType)
                        Exit Sub
                    End If
                End If

                ' 4.b) Altfel, folosește ținta directă a lui "With ..."
                ' găsește primul token după "With" care nu e gol
                Dim baseToken As Token = withTokens.
                                        SkipWhile(Function(t) Not t.TokenString.Equals("With", StringComparison.OrdinalIgnoreCase)).
                                        Skip(1).
                                        FirstOrDefault(Function(t) Not String.IsNullOrWhiteSpace(t.TokenString))

                ' fallback: primul token ne-blank din linie (în caz că n-a prins)
                If baseToken Is Nothing Then
                    baseToken = withTokens.FirstOrDefault(Function(t) Not String.IsNullOrWhiteSpace(t.TokenString) AndAlso
                                              Not t.TokenString.Equals("With", StringComparison.OrdinalIgnoreCase))
                End If

                If baseToken Is Nothing Then Exit Sub

                ' === cazul simplu: With Me
                If baseToken.TokenString.Equals("Me", StringComparison.OrdinalIgnoreCase) Then
                    Dim memberName = tok.TokenString.TrimStart("."c)
                    Dim childSym = TryFindMember(currentContainer, memberName)

                    If childSym IsNot Nothing Then
                        tok.IsResolved = True
                        tok.ResolvedScope = "with_member"
                        tok.ResolvedRef = childSym
                        tok.DataType = childSym.TypeName
                        tok.TokenType = If(childSym.DeclType Is Nothing, childSym.Category, childSym.DeclType)
                        Exit Sub
                    End If
                End If

                ' === cazul general: With <obiect> / With Me.Recordset etc.
                If baseToken.IsResolved AndAlso baseToken.ResolvedRef IsNot Nothing Then
                    Dim parentObj = unwrap(baseToken.ResolvedRef)
                    Dim memberName = tok.TokenString.TrimStart("."c)
                    Dim childSym = TryFindMember(parentObj, memberName)

                    If childSym IsNot Nothing Then
                        tok.IsResolved = True
                        tok.ResolvedScope = "with_member"
                        tok.ResolvedRef = childSym
                        tok.DataType = childSym.TypeName
                        Exit Sub
                    End If
                End If
            End If

            ' === 5️⃣ Parent chaining (fără With) ===
            If tok.Parent IsNot Nothing AndAlso
               tok.Parent.IsResolved AndAlso
               tok.Parent.ResolvedRef IsNot Nothing Then

                ' dacă părintele este un SymbolEntry, extragem obiectul real
                Dim parentObj As Object = tok.Parent.ResolvedRef
                If TypeOf parentObj Is SymbolEntry Then
                    parentObj = DirectCast(parentObj, SymbolEntry).SourceObject
                End If

                Dim memberName = tok.TokenString.TrimStart("."c)
                Dim childSym = TryFindMember(parentObj, memberName)

                If childSym IsNot Nothing Then
                    tok.IsResolved = True
                    tok.ResolvedScope = "member_chain"
                    tok.ResolvedRef = childSym
                    tok.DataType = childSym.TypeName
                    tok.TokenType = If(tok.TokenType = "unresolved", childSym.DeclType, tok.TokenType)
                    Exit Sub
                End If
            End If

            ' === 5️⃣ Module-level symbols ===
            If currentContainer?.ModuleSymbols IsNot Nothing Then
                Dim sym = currentContainer.ModuleSymbols.FirstOrDefault(Function(s) s.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                If sym IsNot Nothing Then
                    tok.IsResolved = True
                    tok.TokenType = sym.DeclType.ToLower()
                    If tok.SourceLine.StartsMethodBlock Or tok.SourceLine.StartsEnumBlock Or tok.SourceLine.StartsTypeBlock Then
                        tok.TokenType &= "_decl"
                    End If
                    tok.ResolvedScope = "module_symbol"
                    tok.ResolvedRef = sym.SourceObject
                    tok.DataType = sym.TypeName
                    Exit Sub
                End If
            End If

            ' === 6️⃣ Unresolved ===
            tok.ResolvedScope = "unresolved"
            tok.IsResolved = False

            If tok.IsResolved Then
                LogInfoLocal($"Resolved '{tok.TokenString}' → {tok.ResolvedScope} ({tok.DataType})", 2)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Function TryFindMember(obj As Object, memberName As String) As SymbolEntry
        If obj Is Nothing OrElse String.IsNullOrEmpty(memberName) Then Return Nothing

        ' === 1️⃣ ModuleContainer ===
        If obj.GetType() Is GetType(ModuleContainer) Then
            Dim modObj = DirectCast(obj, ModuleContainer)
            Dim sym = modObj.ModuleSymbols?.
            FirstOrDefault(Function(s) s.Name.Equals(memberName, StringComparison.OrdinalIgnoreCase))
            If sym IsNot Nothing Then Return sym
        End If

        ' === 2️⃣ FormReportContainer (include Controls + form properties) ===
        If TypeOf obj Is FormReportContainer Then
            Dim frm = DirectCast(obj, FormReportContainer)

            ' (a) Proprietăți de form
            If TryIsKnownFormProperty(memberName) Then
                Return New SymbolEntry With {
                .Name = memberName,
                .TypeName = "Variant",
                .Category = "form_property",
                .SourceObject = frm
            }
            End If

            ' (b) Membri normali declarați în modulul formului
            Dim sym = frm.ModuleSymbols?.
            FirstOrDefault(Function(s) s.Name.Equals(memberName, StringComparison.OrdinalIgnoreCase))
            If sym IsNot Nothing Then Return sym

            ' (c) Controale -- nu cred ca mai e necesar. In frm.ModuleSymbols sunt incarcate si controalele
            Dim ctrl = frm.Controls?.
            FirstOrDefault(Function(c) c.Name.Equals(memberName, StringComparison.OrdinalIgnoreCase))
            If ctrl IsNot Nothing Then
                Return New SymbolEntry With {
                .Name = ctrl.Name,
                .TypeName = ctrl.Type,
                .Category = "control",
                .SourceObject = ctrl
            }
            End If
        End If

        ' === 3️⃣ FormControl (folosește catalogul de proprietăți implicite Access) ===
        If TypeOf obj Is FormControl Then
            Dim ctrl = DirectCast(obj, FormControl)

            ' Proprietăți implicite
            If modAccessPropertyCatalog.TryIsKnownControlProperty(ctrl.Type, memberName) Then
                Return New SymbolEntry With {
                .Name = memberName,
                .TypeName = "Variant",
                .Category = "control_property",
                .SourceObject = ctrl
            }
            End If
        End If

        ' === 4️⃣ Alte obiecte .NET (PropertyInfo, FunctionInfo etc.) ===
        Try
            Dim t = obj.GetType()
            Dim prop = t.GetProperty(memberName)
            If prop IsNot Nothing Then
                Return New SymbolEntry With {
                .Name = memberName,
                .TypeName = prop.PropertyType.Name,
                .Category = "reflection_property",
                .SourceObject = prop
            }
            End If
        Catch
        End Try

        Return Nothing
    End Function

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
                parentLine.Tokens.
                    SkipWhile(Function(t) Not t.TokenString.Equals("With", StringComparison.OrdinalIgnoreCase)).
                    Skip(1).
                    FirstOrDefault(Function(t) Not String.IsNullOrWhiteSpace(t.TokenString))

                If baseTok Is Nothing Then
                    baseTok = parentLine.Tokens.
                    FirstOrDefault(Function(tt) Not String.IsNullOrWhiteSpace(tt.TokenString) AndAlso
                                              Not tt.TokenString.Equals("With", StringComparison.OrdinalIgnoreCase))
                End If

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

    Private Sub LogInfoLocal(ByVal msg As String, level As Integer, Optional force As Boolean = False)
        If pCurrentMethod.Value <> "" Then
            msg = $"({pCurrentType.Value}) [{pCurrentModule.Value}].[{pCurrentMethod.Value}] : {msg}"
        Else
            msg = $"({pCurrentType.Value}) [{pCurrentModule.Value}] : {msg}"
        End If

        LogInfo(msg, level, force)
    End Sub

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
            parentLine.Tokens.
                SkipWhile(Function(t) Not t.TokenString.Equals("With", StringComparison.OrdinalIgnoreCase)).
                Skip(1).
                FirstOrDefault(Function(t) Not String.IsNullOrWhiteSpace(t.TokenString))

            If baseTok Is Nothing Then
                baseTok = parentLine.Tokens.
                FirstOrDefault(Function(tt) Not String.IsNullOrWhiteSpace(tt.TokenString) AndAlso
                                           Not tt.TokenString.Equals("With", StringComparison.OrdinalIgnoreCase))
            End If

            If baseTok IsNot Nothing Then
                Return GetRootToken(baseTok) ' 🔁 urcă recursiv
            End If
        End If

        Return current
    End Function

End Module