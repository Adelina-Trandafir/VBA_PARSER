'PROJECT NAME: VBA_VIEWER
'FILE DESCRIPTION: Helper functions pentru construirea tree-ului
'PATH: VBA_LAUNCHER/TreeHelpers.vb
Imports System.Reflection
Imports System.Windows.Threading
Imports VBA_CORE
Imports VBA_CORE.customTypes
Imports VBA_MethodInfo = VBA_CORE.customTypes.MethodInfo
Imports VBA_PropertyInfo = VBA_CORE.customTypes.PropertyInfo
Imports VBA_FunctionInfo = VBA_CORE.customTypes.FunctionInfo

Namespace VBA_VIEWER
    Module TreeHelpers
        ' TreeHelpers.vb - adaugă la început
        ' TreeHelpers.vb
        Friend Async Function BuildTreeAsync(mainWindow As MainWindow) As Task
            mainWindow.treeMain.Items.Clear()
            mainWindow.moduleInfos = Await Task.Run(AddressOf ExtractModuleInfos)
            mainWindow.formInfos = Await Task.Run(AddressOf ExtractFormInfos)
            Await mainWindow.treeMain.Dispatcher.InvokeAsync(Sub() PopulateTreeView(mainWindow))
        End Function

        Friend Sub PopulateTreeView(mainWindow As MainWindow)
            mainWindow.treeMain.Items.Clear()

            ' 1) Construim toate nodurile principale într-o listă
            Dim topNodes As New List(Of TreeViewItem)

            For Each m In mainWindow.moduleInfos
                Dim node As New TreeViewItem With {.Header = $"📦 {m.Name} ({m.Type})", .Tag = m.Tag}
                node.Items.Add("Loading...")
                AddHandler node.Expanded, AddressOf mainWindow.LazyExpandNode
                topNodes.Add(node)
            Next

            For Each f In mainWindow.formInfos
                Dim node As New TreeViewItem With {.Header = $"🧩 {f.Name} ({f.Type})", .Tag = f.Tag}
                node.Items.Add("Loading...")
                AddHandler node.Expanded, AddressOf mainWindow.LazyExpandNode
                topNodes.Add(node)
            Next

            ' 2) Sortăm DOAR nodurile principale (Modules înaintea Forms/Reports), apoi alfabetic după Name
            For Each n In topNodes.
        OrderBy(Function(x) GetTopNodePriority(x)).ThenBy(Function(x) GetTopNodeName(x), StringComparer.OrdinalIgnoreCase)
                mainWindow.treeMain.Items.Add(n)
            Next
        End Sub

        ' === Helpers pentru sortare la nivel de top ===
        Private Function GetTopNodePriority(n As TreeViewItem) As Integer
            ' Modules/ClassModules primele, apoi Forms/Reports (poți inversa dacă preferi)
            If TypeOf n.Tag Is ModuleContainer Then Return 0
            If TypeOf n.Tag Is FormReportContainer Then Return 1
            Return 2 ' fallback
        End Function

        Private Function GetTopNodeName(n As TreeViewItem) As String
            If TypeOf n.Tag Is ModuleContainer Then
                Return DirectCast(n.Tag, ModuleContainer).Name
            ElseIf TypeOf n.Tag Is FormReportContainer Then
                Return DirectCast(n.Tag, FormReportContainer).Name
            End If
            ' fallback: dacă vreodată Tag nu e setat
            Dim hdr = If(n.Header, "").ToString()
            ' încearcă să decupezi emoji + spațiu, dacă există
            If hdr.Length >= 2 AndAlso Char.IsSurrogatePair(hdr, 0) Then
                Return hdr.Substring(2).Trim()
            End If
            Return hdr
        End Function

        ' ======================================================================
        ' BUILD TREE (complet populat)
        ' ======================================================================
        Private Sub BuildTree(mainWindow As MainWindow)
            mainWindow.treeMain.Items.Clear()

            ' === MODULES ===
            For Each kvp In GlobalModules
                Dim m = kvp.Value
                Dim node = New TreeViewItem With {.Header = $"📦 {m.Name} ({m.Type})", .Tag = m}

                ' Populate module content
                AddModuleContainerNodes(node, m)
                mainWindow.treeMain.Items.Add(node)
                LogInfo($"{m.Name} added")
            Next

            ' === FORMS / REPORTS ===
            For Each kvp In GlobalFormsReports
                Dim f = kvp.Value
                Dim node = New TreeViewItem With {.Header = $"🧩 {f.Name} ({f.Type})", .Tag = f}

                AddFormReportContainerNodes(node, f)
                mainWindow.treeMain.Items.Add(node)
                LogInfo($"{f.Name} added")

            Next
        End Sub

        Friend Sub BuildTreeGroupedByType(ByRef mainWindow As MainWindow, treeMain As TreeView)
            treeMain.Items.Clear()

            Dim nodeClasses = New TreeViewItem With {.Header = "🏗️ Classes", .IsExpanded = False}
            Dim nodeModules = New TreeViewItem With {.Header = "📦 Modules", .IsExpanded = False}
            Dim nodeForms = New TreeViewItem With {.Header = "📝 Forms", .IsExpanded = False}
            Dim nodeReports = New TreeViewItem With {.Header = "📊 Reports", .IsExpanded = False}

            For Each kvp In GlobalModules.OrderBy(Function(k) k.Key)
                Dim m = kvp.Value
                Dim mNode = New TreeViewItem With {.Header = m.Name, .Tag = m}
                AddHandler mNode.Expanded, AddressOf mainWindow.LazyExpandNode
                mNode.Items.Add("Loading...")
                If m.Type = "Module" Then
                    nodeModules.Items.Add(mNode)
                Else
                    nodeClasses.Items.Add(mNode)
                End If
            Next

            For Each kvp In GlobalFormsReports.OrderBy(Function(k) k.Key)
                Dim f = kvp.Value
                Dim fNode = New TreeViewItem With {.Header = f.Name, .Tag = f}
                AddHandler fNode.Expanded, AddressOf mainWindow.LazyExpandNode
                fNode.Items.Add("Loading...")
                If f.Type = "Form" Then nodeForms.Items.Add(fNode) Else nodeReports.Items.Add(fNode)
            Next

            If nodeClasses.Items.Count > 0 Then treeMain.Items.Add(nodeClasses)
            If nodeModules.Items.Count > 0 Then treeMain.Items.Add(nodeModules)
            If nodeForms.Items.Count > 0 Then treeMain.Items.Add(nodeForms)
            If nodeReports.Items.Count > 0 Then treeMain.Items.Add(nodeReports)
        End Sub

        Friend Function ExtractModuleInfos() As List(Of Object)
            Return GlobalModules.Values.Select(AddressOf CreateModuleInfo).ToList()
        End Function

        Friend Function ExtractFormInfos() As List(Of Object)
            Return GlobalFormsReports.Values.Select(AddressOf CreateFormInfo).ToList()
        End Function

        Private Function CreateModuleInfo(m As ModuleContainer) As Object
            Return New With {m.Name, m.Type, .Tag = m}
        End Function

        Private Function CreateFormInfo(f As FormReportContainer) As Object
            Return New With {f.Name, f.Type, .Tag = f}
        End Function

        ' ======================================================================
        ' MODULE CONTAINER STRUCTURE
        ' ======================================================================
        Friend Sub AddModuleContainerNodes(parent As TreeViewItem, m As ModuleContainer)
            AddDeclaresNode(parent, m.Declares)
            AddParamNodes(parent, "📄 Variables", m.Variables)
            AddParamNodes(parent, "📜 Constants", m.Constants)
            AddEnumOrTypeNodes(parent, "🧩 Types", m.Types)
            AddEnumOrTypeNodes(parent, "📚 Enums", m.Enums)
            AddEventsNode(parent, "⚡ Events", m.Events)
            AddPropertyNode(parent, "🧩 Properties", m.Properties)
            AddFunctionsNode(parent, "🔧 Functions", m.Functions)
            AddMethodsNode(parent, m.Methods)
            AddLinesNode(parent, m.Lines)
            AddTokensNode(parent, Nothing)
        End Sub

        ' ======================================================================
        ' FORM / REPORT STRUCTURE
        ' ======================================================================
        Friend Sub AddFormReportContainerNodes(parent As TreeViewItem, f As FormReportContainer)
            AddDeclaresNode(parent, f.Declares)
            AddParamNodes(parent, "📄 Variables", f.Variables)
            AddParamNodes(parent, "📜 Constants", f.Constants)
            AddEnumOrTypeNodes(parent, "🧩 Types", f.Types)
            AddEnumOrTypeNodes(parent, "📚 Enums", f.Enums)
            AddEventsNode(parent, "⚡ Events", f.Events)
            AddPropertyNode(parent, "🧩 Properties", f.Properties)
            AddFunctionsNode(parent, "🔧 Functions", f.Functions)
            AddMethodsNode(parent, f.Methods)
            AddLinesNode(parent, f.Lines)
            AddTokensNode(parent, Nothing)

            ' === Sections ===
            If f.Sections IsNot Nothing AndAlso f.Sections.Count > 0 Then
                Dim secNode = New TreeViewItem With {.Header = "📚 Sections"}
                For Each s In f.Sections
                    Dim sNode = New TreeViewItem With {.Header = $"📖 {s.Type} - {s.Name}", .Tag = s}
                    Dim ctrlNode = New TreeViewItem With {.Header = "🎛 Controls"}
                    For Each c In s.Controls
                        ctrlNode.Items.Add(New TreeViewItem With {.Header = $"🎚 {c.Type}: {c.Name}", .Tag = c})
                    Next
                    sNode.Items.Add(ctrlNode)
                    secNode.Items.Add(sNode)
                Next
                parent.Items.Add(secNode)
            End If

            ' === Controls ===
            If f.Controls IsNot Nothing AndAlso f.Controls.Count > 0 Then
                Dim ctrlRoot = New TreeViewItem With {.Header = "🎛 Controls"}
                For Each c In f.Controls
                    ctrlRoot.Items.Add(New TreeViewItem With {.Header = $"🎚 {c.Type}: {c.Name}", .Tag = c})
                Next
                parent.Items.Add(ctrlRoot)
            End If
        End Sub

        ' ======================================================================
        ' DECLARES NODE
        ' ======================================================================
        Friend Sub AddDeclaresNode(parent As TreeViewItem, declares As Dictionary(Of String, VBA_MethodInfo))
            If declares Is Nothing OrElse declares.Count = 0 Then Return
            Dim dNode = New TreeViewItem With {.Header = "⚙️ Declares", .Tag = declares}
            For Each kvp In declares
                Dim m = kvp.Value
                Dim mNode = New TreeViewItem With {.Header = $"🔹 {m.Name}", .Tag = m}
                AddTokensNode(mNode, m.ParsedLine?.Tokens)
                dNode.Items.Add(mNode)
            Next
            parent.Items.Add(dNode)
        End Sub

        ' ======================================================================
        ' VARIABLES / CONSTANTS
        ' ======================================================================
        Friend Sub AddParamNodes(parent As TreeViewItem, header As String, list As List(Of ParamInfo))
            If list Is Nothing OrElse list.Count = 0 Then Return
            Dim pNode = New TreeViewItem With {.Header = header, .Tag = list}
            For Each p In list
                Dim vNode = New TreeViewItem With {.Header = $"🔹 {p.Name} : {p.Type}", .Tag = p}
                AddTokensNode(vNode, p.Tokens)
                pNode.Items.Add(vNode)
            Next
            parent.Items.Add(pNode)
        End Sub

        ' ======================================================================
        ' VARIABLES / CONSTANTS
        ' ======================================================================
        Friend Sub AddEnumOrTypeNodes(parent As TreeViewItem, header As String, list As List(Of EnumOrType))
            If list Is Nothing OrElse list.Count = 0 Then Return
            Dim pNode = New TreeViewItem With {.Header = header, .Tag = list}
            For Each et In list
                Dim sNode = New TreeViewItem With {.Header = $"🔸 {et.Name} ({et.Type})", .Tag = et}
                For Each p In et.Members
                    Dim vNode = New TreeViewItem With {.Header = $"🔹 {p.Name} : {p.Type}", .Tag = p}
                    AddTokensNode(vNode, p.Tokens)
                    sNode.Items.Add(vNode)
                Next
                pNode.Items.Add(sNode)
            Next
            parent.Items.Add(pNode)
        End Sub

        ' ======================================================================
        ' EVENTS NODE
        ' ======================================================================
        Friend Sub AddEventsNode(parent As TreeViewItem, header As String, events As Dictionary(Of String, VBA_MethodInfo))
            If events Is Nothing OrElse events.Count = 0 Then Return
            Dim eNode = New TreeViewItem With {.Header = "⚡ Events", .Tag = events}
            For Each kvp In events
                Dim m = kvp.Value
                Dim mNode = New TreeViewItem With {.Header = $"🔹 {kvp.Key}", .Tag = m}
                AddTokensNode(mNode, m.Tokens)
                eNode.Items.Add(mNode)
            Next
            parent.Items.Add(eNode)
        End Sub

        ' ======================================================================
        ' METHODS
        ' ======================================================================
        Friend Sub AddMethodsNode(parent As TreeViewItem, methods As Dictionary(Of String, VBA_MethodInfo))
            If methods Is Nothing OrElse methods.Count = 0 Then Return
            Dim mNode = New TreeViewItem With {.Header = "🧠 Methods", .Tag = methods}

            For Each kvp In methods
                Dim method = kvp.Value

                ' Exclude funcțiile și proprietățile (acum sunt în noduri separate)
                If TypeOf method Is VBA_FunctionInfo OrElse TypeOf method Is VBA_PropertyInfo Then Continue For

                Dim mSub = New TreeViewItem With {.Header = $"🔸 {method.Name}", .Tag = method}

                ' Tokens (header line)
                AddTokensNode(mSub, method.ParsedLine?.Tokens)

                ' Parameters
                AddParamNodes(mSub, "🧩 Parameters", method.Parameters)

                ' Variables
                AddParamNodes(mSub, "📄 Variables", method.Variables)

                ' Constants
                AddParamNodes(mSub, "📜 Constants", method.Constants)

                ' MethodLines
                If method.MethodLines IsNot Nothing AndAlso method.MethodLines.Count > 0 Then
                    Dim linesNode = New TreeViewItem With {.Header = "📑 MethodLines", .Tag = method.MethodLines}
                    For Each ln In method.MethodLines
                        Dim lNode = New TreeViewItem With {.Header = $"Line {ln.LineNumber}: {ln.Content}", .Tag = ln}
                        AddTokensNode(lNode, ln.Tokens)
                        linesNode.Items.Add(lNode)
                    Next
                    mSub.Items.Add(linesNode)
                End If

                mNode.Items.Add(mSub)
            Next

            parent.Items.Add(mNode)
        End Sub

        Friend Sub AddPropertyNode(parent As TreeViewItem, header As String, props As Dictionary(Of String, VBA_PropertyInfo))
            If props Is Nothing OrElse props.Count = 0 Then Return
            Dim pNode = New TreeViewItem With {.Header = header, .Tag = props}

            For Each kvp In props
                Dim prop = kvp.Value
                Dim prNode = New TreeViewItem With {.Header = $"🧩 {prop.Name}", .Tag = prop}

                ' Property Type (Get / Let / Set)
                If Not String.IsNullOrEmpty(prop.PropertyType) Then
                    prNode.Items.Add(New TreeViewItem With {.Header = $"🔹 Property Type: {prop.PropertyType}"})
                End If

                ' Return Type
                If Not String.IsNullOrEmpty(prop.ReturnType) Then
                    prNode.Items.Add(New TreeViewItem With {.Header = $"🔹 Return Type: {prop.ReturnType}"})
                End If

                ' Tokens (header line)
                AddTokensNode(prNode, prop.ParsedLine?.Tokens)

                ' Parameters
                AddParamNodes(prNode, "🧩 Parameters", prop.Parameters)

                ' Variables
                AddParamNodes(prNode, "📄 Variables", prop.Variables)

                ' Constants
                AddParamNodes(prNode, "📜 Constants", prop.Constants)

                ' MethodLines
                If prop.MethodLines IsNot Nothing AndAlso prop.MethodLines.Count > 0 Then
                    Dim linesNode = New TreeViewItem With {.Header = "📑 MethodLines", .Tag = prop.MethodLines}
                    For Each ln In prop.MethodLines
                        Dim lNode = New TreeViewItem With {.Header = $"Line {ln.LineNumber}: {ln.Content}", .Tag = ln}
                        AddTokensNode(lNode, ln.Tokens)
                        linesNode.Items.Add(lNode)
                    Next
                    prNode.Items.Add(linesNode)
                End If

                pNode.Items.Add(prNode)
            Next

            parent.Items.Add(pNode)
        End Sub

        Friend Sub AddFunctionsNode(parent As TreeViewItem, header As String, funcs As Dictionary(Of String, FunctionInfo))
            If funcs Is Nothing OrElse funcs.Count = 0 Then Return
            Dim fNode = New TreeViewItem With {.Header = header, .Tag = funcs}

            For Each kvp In funcs
                Dim func = kvp.Value
                Dim fnNode = New TreeViewItem With {.Header = $"🔧 {func.Name}", .Tag = func}

                ' Return Type
                If Not String.IsNullOrEmpty(func.ReturnType) Then
                    fnNode.Items.Add(New TreeViewItem With {.Header = $"🔹 Return Type: {func.ReturnType}"})
                End If

                ' Tokens (header line)
                AddTokensNode(fnNode, func.ParsedLine?.Tokens)

                ' Parameters
                AddParamNodes(fnNode, "🧩 Parameters", func.Parameters)

                ' Variables
                AddParamNodes(fnNode, "📄 Variables", func.Variables)

                ' Constants
                AddParamNodes(fnNode, "📜 Constants", func.Constants)

                ' MethodLines
                If func.MethodLines IsNot Nothing AndAlso func.MethodLines.Count > 0 Then
                    Dim linesNode = New TreeViewItem With {.Header = "📑 MethodLines", .Tag = func.MethodLines}
                    For Each ln In func.MethodLines
                        Dim lNode = New TreeViewItem With {.Header = $"Line {ln.LineNumber}: {ln.Content}", .Tag = ln}
                        AddTokensNode(lNode, ln.Tokens)
                        linesNode.Items.Add(lNode)
                    Next
                    fnNode.Items.Add(linesNode)
                End If

                fNode.Items.Add(fnNode)
            Next

            parent.Items.Add(fNode)
        End Sub

        ' ======================================================================
        ' LINES NODE
        ' ======================================================================
        Friend Sub AddLinesNode(parent As TreeViewItem, lines As List(Of MethodLine))
            If lines Is Nothing OrElse lines.Count = 0 Then Return
            Dim node = New TreeViewItem With {.Header = "🪜 Lines", .Tag = lines}
            For Each ln In lines
                Dim lNode = New TreeViewItem With {.Header = $"Line {ln.LineNumber}: {ln.Content}", .Tag = ln}
                AddTokensNode(lNode, ln.Tokens)
                node.Items.Add(lNode)
            Next
            parent.Items.Add(node)
        End Sub

        ' ======================================================================
        ' TOKENS NODE
        ' ======================================================================
        Friend Sub AddTokensNode(parent As TreeViewItem, tokens As List(Of Token))
            If tokens Is Nothing OrElse tokens.Count = 0 Then Return
            'Dim tNode = New TreeViewItem With {.Header = "💬 Tokens", .Tag = tokens}
            For Each t In tokens
                parent.Items.Add(New TreeViewItem With {
                                .Header = $"💬 {t.TokenString}:{t.TokenType}",
                                .Tag = t})
            Next
        End Sub

        Friend Sub PopulateModuleNode(node As TreeViewItem, m As ModuleContainer)
            If m.Events?.Count > 0 Then AddEventsNode(node, "⚡ Events", m.Events)
            If m.Declares?.Count > 0 Then AddDeclaresNode(node, m.Declares)
            If m.Variables?.Count > 0 Then AddParamNodes(node, "📄 Variables", m.Variables)
            If m.Constants?.Count > 0 Then AddParamNodes(node, "📜 Constants", m.Constants)
            If m.Types?.Count > 0 Then AddEnumOrTypeNodes(node, "🧩 Types", m.Types)
            If m.Enums?.Count > 0 Then AddEnumOrTypeNodes(node, "📚 Enums", m.Enums)

            If m.Properties?.Count > 0 Then AddPropertyNode(node, "🧩 Properties", m.Properties)
            If m.Functions?.Count > 0 Then AddFunctionsNode(node, "🔧 Functions", m.Functions)

            If m.Methods?.Count > 0 Then AddMethodsNode(node, m.Methods)
            If m.Lines?.Count > 0 Then AddLinesNode(node, m.Lines)
        End Sub


        Friend Sub PopulateFormReportNode(node As TreeViewItem, f As FormReportContainer)
            If f.Events?.Count > 0 Then AddEventsNode(node, "⚡ Events", f.Events)
            If f.Declares?.Count > 0 Then AddDeclaresNode(node, f.Declares)
            If f.Variables?.Count > 0 Then AddParamNodes(node, "📄 Variables", f.Variables)
            If f.Constants?.Count > 0 Then AddParamNodes(node, "📜 Constants", f.Constants)

            If f.Types?.Count > 0 Then AddEnumOrTypeNodes(node, "🧩 Types", f.Types)
            If f.Enums?.Count > 0 Then AddEnumOrTypeNodes(node, "📚 Enums", f.Enums)

            If f.Properties?.Count > 0 Then AddPropertyNode(node, "🧩 Properties", f.Properties)
            If f.Functions?.Count > 0 Then AddFunctionsNode(node, "🔧 Functions", f.Functions)

            If f.Methods?.Count > 0 Then AddMethodsNode(node, f.Methods)

            If f.Sections?.Count > 0 Then
                Dim secNode = New TreeViewItem With {.Header = "📚 Sections"}
                For Each s In f.Sections
                    Dim sNode = New TreeViewItem With {.Header = $"📖 {s.Type} - {s.Name}", .Tag = s}
                    secNode.Items.Add(sNode)
                Next
                node.Items.Add(secNode)
            End If

            If f.Controls?.Count > 0 Then
                Dim ctrlNode = New TreeViewItem With {.Header = "🎛 Controls"}
                For Each c In f.Controls
                    ctrlNode.Items.Add(New TreeViewItem With {.Header = $"🎚 {c.Type}: {c.Name}", .Tag = c})
                Next
                node.Items.Add(ctrlNode)
            End If

            If f.Lines?.Count > 0 Then AddLinesNode(node, f.Lines)
        End Sub

        Friend Function CollectChildData(tagObject As Object) As List(Of (Header As String, Tag As Object, Loader As Action(Of TreeViewItem)))
            Dim lst As New List(Of (Header As String, Tag As Object, Loader As Action(Of TreeViewItem)))

            If TypeOf tagObject Is ModuleContainer Then
                CollectModuleData(lst, CType(tagObject, ModuleContainer))
            ElseIf TypeOf tagObject Is FormReportContainer Then
                CollectFormReportData(lst, CType(tagObject, FormReportContainer))
            End If

            Return lst
        End Function

        Private Sub CollectModuleData(lst As List(Of (Header As String, Tag As Object, Loader As Action(Of TreeViewItem))), m As ModuleContainer)
            If m.Declares?.Count > 0 Then
                lst.Add(("⚙️ Declares", m.Declares, Sub(p) AddDeclaresNode(p, m.Declares)))
            End If
            If m.Variables?.Count > 0 Then
                lst.Add(("📄 Variables", m.Variables, Sub(p) AddParamNodes(p, "📄 Variables", m.Variables)))
            End If
            If m.Constants?.Count > 0 Then
                lst.Add(("📜 Constants", m.Constants, Sub(p) AddParamNodes(p, "📜 Constants", m.Constants)))
            End If
            If m.Events?.Count > 0 Then
                lst.Add(("⚡ Events", m.Events, Sub(p) AddEventsNode(p, "⚡ Events", m.Events)))
            End If

            ' ✅ ADĂUGAT — TIPURI ȘI ENUMS
            If m.Types?.Count > 0 Then
                lst.Add(("🧩 Types", m.Types, Sub(p) AddEnumOrTypeNodes(p, "🧩 Types", m.Types)))
            End If

            If m.Enums?.Count > 0 Then
                lst.Add(("📚 Enums", m.Enums, Sub(p) AddEnumOrTypeNodes(p, "📚 Enums", m.Enums)))
            End If

            ' ✅ ADĂUGAT — PROPRIETĂȚI
            If m.Properties?.Count > 0 Then
                lst.Add(("🧩 Properties", m.Properties, Sub(p) AddPropertyNode(p, "🧩 Properties", m.Properties)))
            End If

            ' ✅ ADĂUGAT — FUNCȚII
            If m.Functions?.Count > 0 Then
                lst.Add(("🔧 Functions", m.Functions, Sub(p) AddFunctionsNode(p, "🔧 Functions", m.Functions)))
            End If

            ' ✅ DOAR SUB-URI
            If m.Methods?.Count > 0 Then
                lst.Add(("🧠 Methods", m.Methods, Sub(p) AddMethodsNode(p, m.Methods)))
            End If

            If m.Lines?.Count > 0 Then
                lst.Add(("🪜 Lines", m.Lines, Sub(p) AddLinesNode(p, m.Lines)))
            End If
        End Sub

        Private Sub CollectFormReportData(lst As List(Of (Header As String, Tag As Object, Loader As Action(Of TreeViewItem))), f As FormReportContainer)
            If f.Declares?.Count > 0 Then
                lst.Add(("⚙️ Declares", f.Declares, Sub(p) AddDeclaresNode(p, f.Declares)))
            End If
            If f.Variables?.Count > 0 Then
                lst.Add(("📄 Variables", f.Variables, Sub(p) AddParamNodes(p, "📄 Variables", f.Variables)))
            End If
            If f.Constants?.Count > 0 Then
                lst.Add(("📜 Constants", f.Constants, Sub(p) AddParamNodes(p, "📜 Constants", f.Constants)))
            End If
            If f.Events?.Count > 0 Then
                lst.Add(("⚡ Events", f.Events, Sub(p) AddEventsNode(p, "⚡ Events", f.Events)))
            End If

            ' ✅ ADĂUGAT — TIPURI ȘI ENUMS
            If f.Types?.Count > 0 Then
                lst.Add(("🧩 Types", f.Types, Sub(p) AddEnumOrTypeNodes(p, "🧩 Types", f.Types)))
            End If

            If f.Enums?.Count > 0 Then
                lst.Add(("📚 Enums", f.Enums, Sub(p) AddEnumOrTypeNodes(p, "📚 Enums", f.Enums)))
            End If

            ' ✅ ADĂUGAT — PROPRIETĂȚI
            If f.Properties?.Count > 0 Then
                lst.Add(("🧩 Properties", f.Properties, Sub(p) AddPropertyNode(p, "🧩 Properties", f.Properties)))
            End If

            ' ✅ ADĂUGAT — FUNCȚII
            If f.Functions?.Count > 0 Then
                lst.Add(("🔧 Functions", f.Functions, Sub(p) AddFunctionsNode(p, "🔧 Functions", f.Functions)))
            End If

            ' ✅ DOAR SUB-URI
            If f.Methods?.Count > 0 Then
                lst.Add(("🧠 Methods", f.Methods, Sub(p) AddMethodsNode(p, f.Methods)))
            End If

            If f.Sections?.Count > 0 OrElse f.Controls?.Count > 0 Then
                lst.Add(("📚 Sections & Controls", f, Sub(p) AddFormReportContainerNodes(p, f)))
            End If

            If f.Lines?.Count > 0 Then
                lst.Add(("🪜 Lines", f.Lines, Sub(p) AddLinesNode(p, f.Lines)))
            End If
        End Sub

        ' În MainWindow.xaml.vb, adaugă aceste metode noi:
        Friend Function IsNodeLoading(node As TreeViewItem) As Boolean
            If node.Items.Count <> 1 Then Return False

            If TypeOf node.Items(0) Is String Then
                Return CStr(node.Items(0)).IndexOf("Loading", StringComparison.OrdinalIgnoreCase) >= 0
            ElseIf TypeOf node.Items(0) Is TreeViewItem Then
                Dim ti = DirectCast(node.Items(0), TreeViewItem)
                Dim hdr = If(ti.Header, "").ToString()
                Return hdr.IndexOf("Se încarcă", StringComparison.OrdinalIgnoreCase) >= 0
            End If

            Return False
        End Function

        Friend Sub ShowLoadingState(node As TreeViewItem)
            node.Items.Clear()
            node.Items.Add(New TreeViewItem With {.Header = "⏳ Se încarcă..."})
            node.IsExpanded = True
        End Sub

        Friend Sub PopulateNode(node As TreeViewItem, tagObject As Object, childData As List(Of (Header As String, Tag As Object, Loader As Action(Of TreeViewItem))), loadEx As Exception)
            node.Items.Clear()

            If loadEx IsNot Nothing Then
                node.Items.Add(New TreeViewItem With {.Header = "❌ Eroare: " & loadEx.Message})
                Return
            End If

            If childData Is Nothing OrElse childData.Count = 0 Then
                node.Items.Add(New TreeViewItem With {.Header = "(gol)"})
                Return
            End If

            ' DIRECT adaugă copiii
            If TypeOf tagObject Is ModuleContainer Then
                TreeHelpers.PopulateModuleNode(node, CType(tagObject, ModuleContainer))
            ElseIf TypeOf tagObject Is FormReportContainer Then
                TreeHelpers.PopulateFormReportNode(node, CType(tagObject, FormReportContainer))
            End If
        End Sub

        Friend Sub HighlightTreeNode(ByRef mainWindow As MainWindow, ByRef treeMain As TreeView, moduleName As String, methodName As String)
            For Each n As TreeViewItem In treeMain.Items
                Dim m = TryCast(n.Tag, ModuleContainer)
                Dim f = TryCast(n.Tag, FormReportContainer)

                Dim name = If(m?.Name, f?.Name)
                If String.Equals(name, moduleName, StringComparison.OrdinalIgnoreCase) Then
                    ' expandează nodul modulului
                    n.IsExpanded = True
                    If n.Items.Count > 0 Then mainWindow.LazyExpandNode(n, Nothing)

                    ' caută metoda în ramura Methods
                    For Each subNode As TreeViewItem In n.Items
                        If subNode.Header.ToString().Contains("Methods") Then
                            If subNode.Items.Count = 0 Then mainWindow.LazyExpandNode(subNode, Nothing)
                            For Each methNode As TreeViewItem In subNode.Items
                                If methNode.Header.ToString().IndexOf(methodName, StringComparison.OrdinalIgnoreCase) >= 0 Then
                                    methNode.IsSelected = True
                                    methNode.BringIntoView()
                                    Exit Sub
                                End If
                            Next
                        End If
                    Next
                End If
            Next
        End Sub

        Friend Sub ShowTreeTokenPropertiesPopup(obj As Object, title As String)
            Dim popup As New Window()
            popup.Title = title
            popup.Width = 800
            popup.Height = 650
            popup.WindowStartupLocation = WindowStartupLocation.CenterScreen

            Dim tree As New TreeView() With {
                .FontFamily = New FontFamily("Consolas"),
                .FontSize = 12,
                .Background = New SolidColorBrush(Color.FromRgb(30, 30, 30)),
                .Foreground = New SolidColorBrush(Colors.White),
                .Padding = New Thickness(10)
            }

            ' Generează tree structure - NIVEL 4
            Dim rootNode = DebugExtensionsV3.DumpToTree(obj, 4)

            ' Populează TreeView
            PopulateTreeView(tree, rootNode)

            ' Event pentru double-click pe nod - EXPAND OBIECTE COLLAPSED
            AddHandler tree.MouseDoubleClick, Sub(sender, e)
                                                  Dim selectedItem = TryCast(tree.SelectedItem, TreeViewItem)
                                                  If selectedItem Is Nothing OrElse selectedItem.Tag Is Nothing Then Return

                                                  Dim debugNode = TryCast(selectedItem.Tag, DebugNodeV3)
                                                  If debugNode Is Nothing Then Return

                                                  ' 🔹 Dacă e nod collapsed sau alt tip expandabil
                                                  If debugNode.IsExpandable AndAlso debugNode.OriginalObject IsNot Nothing Then
                                                      Dim objName = If(String.IsNullOrEmpty(debugNode.Name), "Object", debugNode.Name)
                                                      Dim titleText = $"🔍 {objName}: {debugNode.Value}"

                                                      ' 🔸 1) Dacă e collapsed -> expandă obiectul original
                                                      If debugNode.NodeType = "collapsed" Then
                                                          ShowTreeTokenPropertiesPopup(debugNode.OriginalObject, titleText)
                                                          e.Handled = True
                                                          Return
                                                      End If

                                                      ' 🔸 2) Dacă e Tokens / Lines / Collection
                                                      If {"array", "dictionary", "object"}.Contains(debugNode.NodeType) Then
                                                          ShowTreeTokenPropertiesPopup(debugNode.OriginalObject, titleText)
                                                          e.Handled = True
                                                          Return
                                                      End If
                                                  End If
                                              End Sub

            Dim statusBar As New TextBlock() With {
                .Text = "💡 Double-click pe noduri cu 🔍 pentru a expanda în fereastră nouă",
                .Foreground = New SolidColorBrush(Color.FromRgb(180, 180, 180)),
                .Margin = New Thickness(10, 5, 10, 5),
                .FontSize = 10
            }

            Dim btnClose As New Button() With {
                .Content = "Close",
                .Width = 80,
                .Height = 30,
                .Margin = New Thickness(10),
                .HorizontalAlignment = HorizontalAlignment.Center
            }

            AddHandler btnClose.Click, Sub() popup.Close()

            Dim grid As New Grid()
            grid.RowDefinitions.Add(New RowDefinition() With {.Height = New GridLength(1, GridUnitType.Star)})
            grid.RowDefinitions.Add(New RowDefinition() With {.Height = GridLength.Auto})
            grid.RowDefinitions.Add(New RowDefinition() With {.Height = GridLength.Auto})

            Grid.SetRow(tree, 0)
            Grid.SetRow(statusBar, 1)
            Grid.SetRow(btnClose, 2)
            grid.Children.Add(tree)
            grid.Children.Add(statusBar)
            grid.Children.Add(btnClose)

            popup.Content = grid
            popup.ShowDialog()
        End Sub

        Private Sub PopulateTreeView(tree As TreeView, rootNode As DebugNodeV3)
            tree.Items.Clear()

            Dim rootItem = CreateTreeViewItem(rootNode)
            tree.Items.Add(rootItem)

            ' Expand primul nivel
            rootItem.IsExpanded = True
        End Sub

        Private Function CreateTreeViewItem(debugNode As DebugNodeV3) As TreeViewItem
            Dim item As New TreeViewItem() With {.Tag = debugNode}

            ' === HEADER GRID ===
            Dim headerGrid As New Grid() With {.Margin = New Thickness(0, 0, 0, 1)}
            headerGrid.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(260)})
            headerGrid.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = GridLength.Auto})

            ' === (1) NAME ===
            Dim nameBlock As New TextBlock() With {
        .Text = GetNodeIcon(debugNode) & " " & debugNode.Name,
        .Foreground = GetNodeColor(debugNode.NodeType),
        .FontFamily = New FontFamily("Consolas"),
        .FontSize = 12,
        .VerticalAlignment = VerticalAlignment.Center
    }

            ' === (2) VALUE + optional link ===
            Dim valuePanel As New StackPanel() With {.Orientation = Orientation.Horizontal}

            Dim valueBlock As New TextBlock() With {
        .Text = debugNode.Value,
        .Foreground = GetValueColor(debugNode.NodeType),
        .FontFamily = New FontFamily("Consolas"),
        .FontSize = 12,
        .VerticalAlignment = VerticalAlignment.Center
    }
            valuePanel.Children.Add(valueBlock)

            ' 🔗 Dacă are obiect legat real, adaugă link [inspect] colorat în funcție de tip
            If debugNode.LinkedObject IsNot Nothing Then
                Dim objTypeName As String = debugNode.LinkedObject.GetType().Name
                Dim colorBrush As Brush
                Dim icon As String = ""

                Select Case True
                    Case objTypeName.Contains("FunctionInfo")
                        colorBrush = New SolidColorBrush(Color.FromRgb(100, 180, 255)) ' 💙 albastru deschis
                        icon = "🔹"
                    Case objTypeName.Contains("PropertyInfo")
                        colorBrush = New SolidColorBrush(Color.FromRgb(120, 230, 150)) ' 💚 verde
                        icon = "🧩"
                    Case objTypeName.Contains("ParamInfo")
                        colorBrush = New SolidColorBrush(Color.FromRgb(255, 190, 100)) ' 🧡 portocaliu
                        icon = "🔸"
                    Case objTypeName.Contains("ModuleContainer")
                        colorBrush = New SolidColorBrush(Color.FromRgb(200, 160, 255)) ' 💜 violet
                        icon = "🧱"
                    Case objTypeName.Contains("MethodInfo")
                        colorBrush = New SolidColorBrush(Color.FromRgb(150, 210, 255)) ' 🔵 albastru
                        icon = "🔷"
                    Case Else
                        colorBrush = New SolidColorBrush(Color.FromRgb(180, 180, 180)) ' ⚪ gri deschis
                        icon = "⚪"
                End Select

                Dim linkText As New TextBlock() With {
            .Text = $" {icon} [inspect]",
            .Foreground = colorBrush,
            .FontStyle = FontStyles.Italic,
            .Cursor = Cursors.Hand,
            .TextDecorations = TextDecorations.Underline,
            .Margin = New Thickness(5, 0, 0, 0),
            .ToolTip = $"Inspect {objTypeName}"
        }

                AddHandler linkText.MouseLeftButtonUp, Sub()
                                                           Dim objName = If(String.IsNullOrEmpty(debugNode.Name), "Object", debugNode.Name)
                                                           Dim titleText = $"🔍 {objName} → {objTypeName}"
                                                           ShowTreeTokenPropertiesPopup(debugNode.LinkedObject, titleText)
                                                       End Sub

                AddHandler linkText.MouseEnter, Sub()
                                                    linkText.Foreground = New SolidColorBrush(Color.FromRgb(255, 255, 255))
                                                End Sub
                AddHandler linkText.MouseLeave, Sub()
                                                    linkText.Foreground = colorBrush
                                                End Sub

                valuePanel.Children.Add(linkText)
            End If

            ' === COMBINARE GRID ===
            Grid.SetColumn(nameBlock, 0)
            Grid.SetColumn(valuePanel, 1)
            headerGrid.Children.Add(nameBlock)
            headerGrid.Children.Add(valuePanel)

            item.Header = headerGrid

            ' === COPII ===
            For Each child In debugNode.Children
                item.Items.Add(CreateTreeViewItem(child))
            Next

            Return item
        End Function


        Private Function GetValueColor(nodeType As String) As Brush
            Select Case nodeType
                Case "property", "object"
                    Return New SolidColorBrush(Color.FromRgb(190, 190, 190))  ' gri deschis
                Case "array", "dictionary"
                    Return New SolidColorBrush(Color.FromRgb(224, 175, 120))  ' bej/portocaliu
                Case "primitive"
                    Return New SolidColorBrush(Color.FromRgb(181, 206, 168))  ' verde deschis
                Case "reference"
                    Return New SolidColorBrush(Color.FromRgb(255, 255, 140))  ' galben pal
                Case "collapsed"
                    Return New SolidColorBrush(Color.FromRgb(140, 220, 255))  ' albastru deschis
                Case Else
                    Return New SolidColorBrush(Colors.White)
            End Select
        End Function

        Private Function GetNodeColor(nodeType As String) As Brush
            Select Case nodeType
                Case "property"
                    Return New SolidColorBrush(Color.FromRgb(156, 220, 254))  ' Albastru
                Case "object"
                    Return New SolidColorBrush(Color.FromRgb(220, 220, 220))  ' Alb
                Case "array", "dictionary"
                    Return New SolidColorBrush(Color.FromRgb(206, 145, 120))  ' Portocaliu
                Case "primitive"
                    Return New SolidColorBrush(Color.FromRgb(181, 206, 168))  ' Verde
                Case "reference"
                    Return New SolidColorBrush(Color.FromRgb(255, 215, 0))    ' Galben
                Case "collapsed"
                    Return New SolidColorBrush(Color.FromRgb(100, 200, 255))  ' Albastru deschis - indica că e clickable
                Case Else
                    Return New SolidColorBrush(Colors.White)
            End Select
        End Function

        Private Function GetNodeIcon(node As DebugNodeV3) As String
            Select Case node.NodeType
                Case "object" : Return "🧩"
                Case "array" : Return "📦"
                Case "dictionary" : Return "📚"
                Case "property" : Return "🔸"
                Case "primitive" : Return "🔹"
                Case "collapsed" : Return "🔍"
                Case "reference" : Return "↻"
                Case Else
                    ' Obiecte rezolvate special (ex: token, method, property)
                    If node.OriginalObject IsNot Nothing Then
                        Dim tName = node.OriginalObject.GetType().Name.ToLowerInvariant()
                        If tName.Contains("token") Then Return "🔖"
                        If tName.Contains("methodinfo") Then Return "🧠"
                        If tName.Contains("propertyinfo") Then Return "🏷️"
                        If tName.Contains("functioninfo") Then Return "🔧"
                        If tName.Contains("paraminfo") Then Return "🧾"
                    End If
                    Return "▫️"
            End Select
        End Function
    End Module
End Namespace