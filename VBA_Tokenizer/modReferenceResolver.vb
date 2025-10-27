'Imports System.Collections.Concurrent
'Imports System.IO
'Imports System.Text.RegularExpressions
'Imports AVACONT_Core.modTypes
'Imports AVACONT_Core
'Imports AVACONT_Core.Logger

'Public Module modReferenceResolver

'    ' === Heuristică events tipice VBA pe Forms/Reports (exclude din dead-code) ===
'    Public Function IsTypicalAccessEvent(funcName As String) As Boolean
'        Dim methodName As String = funcName.Split("."c).Last()
'        Return Regex.IsMatch(methodName, "^(Form_|Report_|On|Before|After|Click|Load|Current|Open|Close|Dirty|Change|Timer|Mouse|Key|Btn_)", RegexOptions.IgnoreCase) _
'            OrElse Regex.IsMatch(methodName, "(_Click|_Change|_Enter|_Exit|_Update|_Load|_Open|_Close)$", RegexOptions.IgnoreCase)
'    End Function

'    ' ===== Utils chei/lookup =====
'    Private Function ModuleKey(em As ExportedModule) As String
'        Return $"{em.Type}:{em.Name}"
'    End Function

'    Private Function ModuleKey(typeName As String, modName As String) As String
'        Return $"{typeName}:{modName}"
'    End Function

'    Private Function PreferOrder(typeName As String) As Integer
'        ' prioritate la dezambiguizare: Modules > Forms > Reports > Macros
'        Select Case True
'            Case typeName.Equals("Modules", StringComparison.OrdinalIgnoreCase) : Return 0
'            Case typeName.Equals("Forms", StringComparison.OrdinalIgnoreCase) : Return 1
'            Case typeName.Equals("Reports", StringComparison.OrdinalIgnoreCase) : Return 2
'            Case typeName.Equals("Macros", StringComparison.OrdinalIgnoreCase) : Return 3
'            Case Else : Return 9
'        End Select
'    End Function

'    Private Function ExtractMember(methodOrQualified As String) As String
'        If String.IsNullOrEmpty(methodOrQualified) Then Return ""
'        Return methodOrQualified.Split("."c).Last()
'    End Function

'    ' Identifică modulul „caller” din lista de candidați cu același Name
'    Private Function PickCallerModule(cands As List(Of ExportedModule), callerFunc As String) As ExportedModule
'        If cands Is Nothing OrElse cands.Count = 0 Then Return Nothing
'        If cands.Count = 1 Then Return cands(0)

'        Dim needleFull As String = callerFunc
'        Dim needleMethod As String = ExtractMember(callerFunc)

'        ' 1) dacă funcția apelantă există în modul
'        For Each em In cands
'            If em.Functions.Any(Function(f) f.Name.Equals(needleFull, StringComparison.OrdinalIgnoreCase) OrElse
'                                           ExtractMember(f.Name).Equals(needleMethod, StringComparison.OrdinalIgnoreCase)) Then
'                Return em
'            End If
'        Next
'        ' 2) fallback: după ordine de preferință a tipului
'        Return cands.OrderBy(Function(x) PreferOrder(x.Type)).First()
'    End Function

'    Public Sub ResolveReferences(exported As List(Of ExportedModule))
'        ' === Indici pentru module ===
'        Dim modByKey As New Dictionary(Of String, ExportedModule)(StringComparer.OrdinalIgnoreCase)
'        Dim modsByName As New Dictionary(Of String, List(Of ExportedModule))(StringComparer.OrdinalIgnoreCase)
'        For Each em In exported
'            Dim key = ModuleKey(em)
'            modByKey(key) = em
'            If Not modsByName.ContainsKey(em.Name) Then modsByName(em.Name) = New List(Of ExportedModule)
'            modsByName(em.Name).Add(em)
'        Next

'        ' === funcIndex pe modul (cheie compusă) → nume metodă simplu → listă funcții ===
'        Dim funcIndex As New Dictionary(Of String, Dictionary(Of String, List(Of FunctionInfo)))(StringComparer.OrdinalIgnoreCase)
'        For Each em In exported
'            Dim dict As New Dictionary(Of String, List(Of FunctionInfo))(StringComparer.OrdinalIgnoreCase)
'            For Each f In em.Functions
'                Dim simpleName = ExtractMember(f.Name)
'                If Not dict.ContainsKey(simpleName) Then dict(simpleName) = New List(Of FunctionInfo)
'                dict(simpleName).Add(f)
'            Next
'            funcIndex(ModuleKey(em)) = dict
'        Next

'        ' === simboluri: nume → toate definițiile (în module diferite) ===
'        Dim symbolIndex = GlobalSymbols.
'            GroupBy(Function(s) s.Name, StringComparer.OrdinalIgnoreCase).
'            ToDictionary(Function(g) g.Key, Function(g) g.ToList(), StringComparer.OrdinalIgnoreCase)

'        ' === varIndex: moduleName (doar nume!) → scopeFunc → varName → TypeName
'        ' Notă: moduleName e doar Name. Dacă există dubluri de Name, dezambiguizăm pe baza caller-ului ales.
'        Dim varIndex As New Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, String)))(StringComparer.OrdinalIgnoreCase)
'        For Each v In GlobalVariables
'            If Not varIndex.ContainsKey(v.CallerModule) Then
'                varIndex(v.CallerModule) = New Dictionary(Of String, Dictionary(Of String, String))(StringComparer.OrdinalIgnoreCase)
'            End If
'            If Not varIndex(v.CallerModule).ContainsKey(v.ScopeFunc) Then
'                varIndex(v.CallerModule)(v.ScopeFunc) = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
'            End If
'            varIndex(v.CallerModule)(v.ScopeFunc)(v.Name) = v.TypeName
'        Next

'        ' lookup tip variabilă din cadrul modulului „ales”
'        Dim resolveVarType As Func(Of ExportedModule, String, String, String) =
'            Function(callerEm As ExportedModule, scopeFunc As String, varName As String) As String
'                If callerEm Is Nothing Then Return ""
'                Dim modName As String = callerEm.Name ' atenție: key-ul pt varIndex e doar Name
'                Dim t As String = ""
'                Dim dictMod As Dictionary(Of String, Dictionary(Of String, String)) = Nothing
'                If varIndex.TryGetValue(modName, dictMod) Then
'                    Dim dictFunc As Dictionary(Of String, String) = Nothing
'                    If Not String.IsNullOrEmpty(scopeFunc) AndAlso dictMod.TryGetValue(scopeFunc, dictFunc) Then
'                        If dictFunc.TryGetValue(varName, t) Then Return t
'                    End If
'                    If dictMod.TryGetValue("", dictFunc) AndAlso dictFunc.TryGetValue(varName, t) Then
'                        Return t
'                    End If
'                End If
'                Return ""
'            End Function

'        Dim refs = GlobalReferences.ToList()
'        Dim linksMade As Integer = 0

'        For Each r In refs
'            ' — găsește modulul caller (poate exista dublură de nume; alegem prin funcția apelantă)
'            Dim callerEm As ExportedModule = Nothing
'            Dim lst As List(Of ExportedModule) = Nothing
'            If modsByName.TryGetValue(r.CallerModule, lst) Then
'                callerEm = PickCallerModule(lst, r.CallerFunc)
'            End If
'            If callerEm Is Nothing Then Continue For

'            Dim callerKey As String = ModuleKey(callerEm)

'            Dim token = r.Token
'            Dim calleeCandidates As New List(Of (calleeMod As ExportedModule, finfo As FunctionInfo))

'            If token.Contains("."c) Then
'                ' pattern: a.Member(
'                Dim parts = token.Split("."c)
'                If parts.Length = 2 Then
'                    Dim varName = parts(0)
'                    Dim memberName = parts(1)
'                    Dim tp = resolveVarType(callerEm, r.CallerFunc, varName)
'                    If String.IsNullOrEmpty(tp) Then GoTo DoneToken

'                    Dim listSyms As List(Of SymbolInfo) = Nothing
'                    If symbolIndex.TryGetValue(memberName, listSyms) Then
'                        For Each s In listSyms
'                            ' DeclaringModule = doar nume; pot exista mai multe module cu acel nume → testează toate
'                            If Not s.DeclaringClass.Equals(tp, StringComparison.OrdinalIgnoreCase) Then Continue For

'                            Dim targets As List(Of ExportedModule) = Nothing
'                            If modsByName.TryGetValue(s.DeclaringModule, targets) Then
'                                For Each em In targets
'                                    Dim emKey = ModuleKey(em)
'                                    Dim dict As Dictionary(Of String, List(Of FunctionInfo)) = Nothing
'                                    If funcIndex.TryGetValue(emKey, dict) Then
'                                        Dim listF As List(Of FunctionInfo) = Nothing
'                                        If dict.TryGetValue(memberName, listF) Then
'                                            For Each f In listF
'                                                ' member definit în em și cu nume compatibil
'                                                calleeCandidates.Add((em, f))
'                                            Next
'                                        End If
'                                    End If
'                                Next
'                            End If
'                        Next
'                    End If
'                End If
'            Else
'                ' pattern: BareCall(
'                Dim memberName = token
'                Dim listSyms As List(Of SymbolInfo) = Nothing
'                If symbolIndex.TryGetValue(memberName, listSyms) Then
'                    For Each s In listSyms
'                        Dim sameModuleName = s.DeclaringModule.Equals(callerEm.Name, StringComparison.OrdinalIgnoreCase)
'                        ' DeclaringModule e doar nume → pot exista mai multe module cu acel nume
'                        Dim targets As List(Of ExportedModule) = Nothing
'                        If modsByName.TryGetValue(s.DeclaringModule, targets) Then
'                            For Each em In targets
'                                Dim sameModule = sameModuleName AndAlso em.Type.Equals(callerEm.Type, StringComparison.OrdinalIgnoreCase)
'                                If sameModule OrElse s.IsPublic Then
'                                    Dim emKey = ModuleKey(em)
'                                    Dim dict As Dictionary(Of String, List(Of FunctionInfo)) = Nothing
'                                    If funcIndex.TryGetValue(emKey, dict) Then
'                                        Dim listF As List(Of FunctionInfo) = Nothing
'                                        If dict.TryGetValue(memberName, listF) Then
'                                            For Each f In listF
'                                                If sameModule OrElse Not f.Accessibility.Equals("Private", StringComparison.OrdinalIgnoreCase) Then
'                                                    calleeCandidates.Add((em, f))
'                                                End If
'                                            Next
'                                        End If
'                                    End If
'                                End If
'                            Next
'                        End If
'                    Next
'                End If
'            End If

'DoneToken:
'            ' leagă caller → toți candidații validați
'            For Each cand In calleeCandidates
'                Dim callee = cand.calleeMod
'                Dim finfo = cand.finfo

'                SyncLock callee
'                    callee.IncomingModules.Add(callerEm.Name)
'                    callerEm.OutgoingModules.Add(callee.Name)

'                    Dim refTag = $"{callerEm.Name}:{r.LineNumber}"
'                    If Not finfo.CalledBy.Contains(refTag) Then finfo.CalledBy.Add(refTag)
'                    If Not callee.References.Contains(refTag) Then callee.References.Add(refTag)
'                End SyncLock

'                linksMade += 1
'            Next
'        Next

'        LogInfo($"[Resolver] Links made: {linksMade}")
'    End Sub

'    ' ===== Maps (din *_maps.txt) – RunCode / simboluri simple =====
'    Public Sub AnalyzeMapsReferences(exported As List(Of ExportedModule))
'        If mapCache.Count = 0 Then Exit Sub

'        ' Indici module
'        Dim modByKey As New Dictionary(Of String, ExportedModule)(StringComparer.OrdinalIgnoreCase)
'        Dim modsByName As New Dictionary(Of String, List(Of ExportedModule))(StringComparer.OrdinalIgnoreCase)
'        For Each em In exported
'            modByKey(ModuleKey(em)) = em
'            If Not modsByName.ContainsKey(em.Name) Then modsByName(em.Name) = New List(Of ExportedModule)
'            modsByName(em.Name).Add(em)
'        Next

'        ' funcții publice indexate după nume (fără clasă)
'        Dim nameToPublicFns As New Dictionary(Of String, List(Of (m As ExportedModule, f As FunctionInfo)))(StringComparer.OrdinalIgnoreCase)
'        For Each em In exported
'            For Each f In em.Functions
'                If f.Accessibility.Equals("Private", StringComparison.OrdinalIgnoreCase) Then Continue For
'                Dim methodOnly = ExtractMember(f.Name)
'                If Not nameToPublicFns.ContainsKey(methodOnly) Then nameToPublicFns(methodOnly) = New List(Of (ExportedModule, FunctionInfo))()
'                nameToPublicFns(methodOnly).Add((em, f))
'            Next
'        Next

'        For Each kv In mapCache
'            Dim mapPath = kv.Key
'            Dim txt = kv.Value

'            ' deduce Type din folderul părinte: ...\ExportVBA\<Type>\File_maps.txt
'            Dim ownerType As String = New DirectoryInfo(Path.GetDirectoryName(mapPath)).Name
'            Dim ownerName As String = Path.GetFileNameWithoutExtension(mapPath)
'            ownerName = Regex.Replace(ownerName, "_maps$", "", RegexOptions.IgnoreCase)

'            Dim owner As ExportedModule = Nothing
'            Dim ownerKey = ModuleKey(ownerType, ownerName)
'            If Not modByKey.TryGetValue(ownerKey, owner) Then
'                ' fallback: dacă tipul din folder nu e standard, încearcă pe nume (poate un singur match)
'                Dim lst As List(Of ExportedModule) = Nothing
'                If modsByName.TryGetValue(ownerName, lst) AndAlso lst.Count = 1 Then
'                    owner = lst(0)
'                Else
'                    Continue For
'                End If
'            End If

'            Dim m1 = Regex.Matches(txt, "RunCode\(\s*""\s*([A-Za-z0-9_\.]+)\s*""\s*\)", RegexOptions.IgnoreCase)
'            Dim m2 = Regex.Matches(txt, "=\s*([A-Za-z_][A-Za-z0-9_]*)(?:\s*\(|\b)", RegexOptions.IgnoreCase)

'            Dim hits As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
'            For Each m In m1.Cast(Of Match)() : hits.Add(m.Groups(1).Value) : Next
'            For Each m In m2.Cast(Of Match)() : hits.Add(m.Groups(1).Value) : Next

'            For Each h In hits
'                Dim funcName = ExtractMember(h)
'                If nameToPublicFns.ContainsKey(funcName) Then
'                    For Each pair In nameToPublicFns(funcName)
'                        Dim callee = pair.m
'                        Dim finfo = pair.f
'                        SyncLock callee
'                            Dim tag = $"{owner.Name}:maps"
'                            If Not callee.References.Contains(tag) Then callee.References.Add(tag)
'                            callee.IncomingModules.Add(owner.Name)
'                            owner.OutgoingModules.Add(callee.Name)
'                            If Not finfo.CalledBy.Contains(tag) Then finfo.CalledBy.Add(tag)
'                        End SyncLock
'                    Next
'                End If
'            Next
'        Next
'    End Sub

'    ' ===== Cicluri =====
'    Public Function DetectModuleCycles(exported As List(Of ExportedModule)) As List(Of String)
'        ' folosește key compus Type:Name
'        Dim keyToModule As Dictionary(Of String, ExportedModule) =
'            exported.ToDictionary(Function(m) ModuleKey(m), Function(m) m, StringComparer.OrdinalIgnoreCase)

'        Dim cycles As New List(Of String)
'        Dim visited As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
'        Dim stack As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
'        Dim path As New List(Of String)

'        Dim neighbors = Function(k As String) As IEnumerable(Of String)
'                            Dim m As ExportedModule = Nothing
'                            If keyToModule.TryGetValue(k, m) Then
'                                ' map Name → chei ale vecinilor (pot exista mai multe cu același nume)
'                                Dim outs = New List(Of String)
'                                For Each outName In m.OutgoingModules
'                                    ' pot exista mai multe module cu acest nume; ia-le pe toate cheile
'                                    For Each kv In keyToModule
'                                        If kv.Value.Name.Equals(outName, StringComparison.OrdinalIgnoreCase) Then
'                                            outs.Add(kv.Key)
'                                        End If
'                                    Next
'                                Next
'                                Return outs
'                            End If
'                            Return Enumerable.Empty(Of String)()
'                        End Function

'        Dim dfs As Action(Of String)
'        dfs =
'            Sub(u As String)
'                visited.Add(u)
'                stack.Add(u)
'                path.Add(u)
'                For Each v In neighbors(u)
'                    If Not visited.Contains(v) Then
'                        dfs(v)
'                    ElseIf stack.Contains(v) Then
'                        Dim idx = path.IndexOf(v)
'                        If idx >= 0 Then
'                            Dim cyc = path.Skip(idx).ToList()
'                            cyc.Add(v)
'                            cycles.Add(String.Join(" -> ", cyc))
'                        End If
'                    End If
'                Next
'                stack.Remove(u)
'                path.RemoveAt(path.Count - 1)
'            End Sub

'        For Each k In keyToModule.Keys
'            If Not visited.Contains(k) Then dfs(k)
'        Next

'        Return cycles.Distinct().ToList()
'    End Function

'End Module
