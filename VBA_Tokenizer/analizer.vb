'Imports AVACONT_Core
'Imports AVACONT_Core.Logger

'Module Analizer_Classifier
'    ' ==============================================================
'    '  CLASIFICARE TOKEN-URI & ANALIZA DEPENDENȚELOR
'    ' ==============================================================
'    Public Sub ClassifyTokensAndAnalyzeDependencies()
'        LogInfo("=== ClassifyTokensAndAnalyzeDependencies ===")

'        ' Resetare liste
'        GlobalLocalSymbols = New Concurrent.ConcurrentBag(Of LocalSymbolInfo)
'        GlobalNonStandardCalls = New Concurrent.ConcurrentBag(Of NonStandardCallInfo)

'        ' === Construim index-uri GLOBALE (o singură dată, înainte de paralelizare) ===

'        ' 1) Index: toate simbolurile GLOBAL (PUBLIC din Module standard)
'        Dim globalPublicSymbols As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
'        For Each sym In GlobalSymbols
'            If sym.IsPublic AndAlso Not IsClassLikeModule(sym.DeclaringModule) Then
'                globalPublicSymbols.Add(sym.Name)
'            End If
'        Next

'        ' 2) Index: controale per formular (thread-safe dictionary)
'        Dim formControlsIndex As New Concurrent.ConcurrentDictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)
'        For Each kvp In GlobalFormsReports
'            Dim controlNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
'            For Each ctrl In kvp.Value.Controls
'                controlNames.Add(ctrl.Name)
'            Next
'            formControlsIndex.TryAdd(kvp.Key, controlNames)
'        Next

'        ' 3) Index: variabile locale per modul+funcție (thread-safe)
'        Dim localVarsIndex As New Concurrent.ConcurrentDictionary(Of String, Dictionary(Of String, HashSet(Of String)))(StringComparer.OrdinalIgnoreCase)
'        For Each varInfo In GlobalVariables
'            Dim modDict = localVarsIndex.GetOrAdd(varInfo.CallerModule, Function(k) New Dictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase))

'            SyncLock modDict
'                If Not modDict.ContainsKey(varInfo.ScopeFunc) Then
'                    modDict(varInfo.ScopeFunc) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
'                End If
'                modDict(varInfo.ScopeFunc).Add(varInfo.Name)
'            End SyncLock

'            ' Adaugă în GlobalLocalSymbols (thread-safe bag)
'            GlobalLocalSymbols.Add(New LocalSymbolInfo With {
'            .Name = varInfo.Name,
'            .DeclaringModule = varInfo.CallerModule,
'            .DeclaringScope = varInfo.ScopeFunc,
'            .SymbolType = "variable",
'            .DataType = varInfo.TypeName,
'            .IsPrivate = True
'        })
'        Next

'        ' 4) Index: funcții/metode private per modul - PARSARE DIN SURSE
'        Dim privateMethodsIndex As New Concurrent.ConcurrentDictionary(Of String, HashSet(Of String))(StringComparer.OrdinalIgnoreCase)

'        ' 4a) Parcurge GlobalFormsReports pentru metode private din Forms/Reports
'        For Each kvp In GlobalFormsReports
'            Dim moduleName = kvp.Key
'            Dim fci = kvp.Value

'            If fci.Methods Is Nothing Then Continue For

'            Dim methodSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

'            For Each mvp In fci.Methods
'                Dim method = mvp.Value
'                If method.MethodScope.Equals("private", StringComparison.OrdinalIgnoreCase) Then
'                    methodSet.Add(method.Name)

'                    ' Adaugă în GlobalLocalSymbols
'                    Dim symType As String = "function"
'                    If method.Signature.IndexOf("Sub ", StringComparison.OrdinalIgnoreCase) >= 0 Then symType = "sub"
'                    If method.Signature.IndexOf("Property ", StringComparison.OrdinalIgnoreCase) >= 0 Then symType = "property"

'                    GlobalLocalSymbols.Add(New LocalSymbolInfo With {
'                    .Name = method.Name,
'                    .DeclaringModule = moduleName,
'                    .DeclaringScope = moduleName,
'                    .SymbolType = symType,
'                    .DataType = "",
'                    .IsPrivate = True
'                })
'                End If
'            Next

'            If methodSet.Count > 0 Then
'                privateMethodsIndex.TryAdd(moduleName, methodSet)
'            End If
'        Next

'        ' 4b) Parcurge GlobalModules pentru metode private din Modules/Classes
'        'For Each gm In GlobalModules
'        '    Dim moduleName = gm.Name
'        '    Dim methodSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

'        '    ' Parsează rapid pentru a găsi metodele private
'        '    Dim lines = gm.CodeContent.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
'        '    For Each line In lines
'        '        Dim m = RegexCache.RxHeader.Match(line)
'        '        If m.Success Then
'        '            Dim access = If(m.Groups(1).Success, m.Groups(1).Value, "Public")
'        '            Dim name = If(m.Groups(4).Success, m.Groups(4).Value, "")

'        '            If access.Equals("Private", StringComparison.OrdinalIgnoreCase) AndAlso name <> "" Then
'        '                methodSet.Add(name)

'        '                ' Determină tipul metodei
'        '                Dim symType As String = "function"
'        '                If m.Groups(2).Success Then
'        '                    Dim kind = m.Groups(2).Value
'        '                    If kind.Equals("Sub", StringComparison.OrdinalIgnoreCase) Then symType = "sub"
'        '                End If
'        '                If m.Groups(3).Success Then symType = "property"

'        '                GlobalLocalSymbols.Add(New LocalSymbolInfo With {
'        '                .Name = name,
'        '                .DeclaringModule = moduleName,
'        '                .DeclaringScope = "",
'        '                .SymbolType = symType,
'        '                .DataType = "",
'        '                .IsPrivate = True
'        '            })
'        '            End If
'        '        End If
'        '    Next

'        '    If methodSet.Count > 0 Then
'        '        privateMethodsIndex.TryAdd(moduleName, methodSet)
'        '    End If
'        'Next

'        ' === Pregătim contextul global pentru procesare ===
'        Dim context As New ClassificationContext With {
'        .globalPublicSymbols = globalPublicSymbols,
'        .formControlsIndex = formControlsIndex,
'        .localVarsIndex = localVarsIndex,
'        .privateMethodsIndex = privateMethodsIndex
'    }

'        ' === PARALELIZARE PE FIECARE FORM/MODULE ===
'        Dim opts As New ParallelOptions With {.MaxDegreeOfParallelism = Environment.ProcessorCount}

'        If useParallel Then
'            ' Procesare paralelă - fiecare formular în thread separat
'            Parallel.ForEach(GlobalFormsReports, opts, Sub(kvp) ProcessSingleModule(kvp.Key, kvp.Value, context))
'        Else
'            ' Procesare secvențială
'            For Each kvp In GlobalFormsReports
'                ProcessSingleModule(kvp.Key, kvp.Value, context)
'            Next
'        End If

'        LogInfo($"[Classifier] LocalSymbols: {GlobalLocalSymbols.Count}, NonStandardCalls: {GlobalNonStandardCalls.Count}")
'    End Sub

'    ' === Procesează UN SINGUR modul/formular (rulează paralel) ===
'    Private Sub ProcessSingleModule(moduleName As String, fci As FormReportContainer, context As ClassificationContext)
'        If fci Is Nothing OrElse fci.Methods Is Nothing Then Return

'        Try
'            Dim isClassLike As Boolean = IsClassLikeModule(moduleName)
'            Dim controlNames As HashSet(Of String) = Nothing
'            context.FormControlsIndex.TryGetValue(moduleName, controlNames)

'            Dim modDict As Dictionary(Of String, HashSet(Of String)) = Nothing
'            context.LocalVarsIndex.TryGetValue(moduleName, modDict)

'            Dim privateMethods As HashSet(Of String) = Nothing
'            context.PrivateMethodsIndex.TryGetValue(moduleName, privateMethods)

'            ' Procesează toate metodele din acest modul
'            For Each kvp In fci.Methods
'                Dim method = kvp.Value
'                If method.Tokens Is Nothing Then Continue For

'                Dim methodFullName = If(isClassLike, $"{moduleName}.{method.Name}", method.Name)

'                ' Obține variabilele locale pentru această metodă
'                Dim localVars As HashSet(Of String) = Nothing
'                If modDict IsNot Nothing AndAlso modDict.ContainsKey(methodFullName) Then
'                    localVars = modDict(methodFullName)
'                End If

'                ' Procesează fiecare token
'                For Each tokenInfo In method.Tokens
'                    Dim token = tokenInfo.TokenString

'                    ' === 1) VERIFICĂ DACĂ E VBA KEYWORD ===
'                    If RegexCache.VbaKeywords.Contains(token) Then Continue For

'                    ' === 2) VERIFICĂ DACĂ E CONTROL ===
'                    If controlNames IsNot Nothing AndAlso controlNames.Contains(token) Then Continue For

'                    ' === 3) VERIFICĂ DACĂ E VARIABILĂ LOCALĂ ===
'                    If localVars IsNot Nothing AndAlso localVars.Contains(token) Then Continue For

'                    ' === 4) VERIFICĂ DACĂ E FUNCȚIE PRIVATĂ DIN ACELAȘI MODUL ===
'                    If privateMethods IsNot Nothing AndAlso privateMethods.Contains(token) Then Continue For

'                    ' === 5) VERIFICĂ DACĂ E GLOBAL PUBLIC (din Module standard) ===
'                    If context.GlobalPublicSymbols.Contains(token) Then Continue For

'                    ' === 6) VERIFICĂ LISTA DE IGNORE (built-in VBA, Me, etc.) ===
'                    If VBAIgnoreList.ShouldIgnore(token, moduleName, "Forms", fci) Then Continue For

'                    ' === 7) E UN APEL NON-STANDARD → Adaugă în GlobalNonStandardCalls ===
'                    GlobalNonStandardCalls.Add(New NonStandardCallInfo With {
'                        .CallerModule = moduleName,
'                        .CallerMethod = methodFullName,
'                        .CalleeToken = token,
'                        .LineNumber = tokenInfo.LineNumber,
'                        .IsQualified = token.Contains("."c)
'                    })
'                Next
'            Next

'        Catch ex As Exception
'            LogError($"ProcessSingleModule({moduleName})", ex)
'        End Try
'    End Sub
'End Module
