Imports System.Text
Imports System.Windows
Imports System.Windows.Controls
Imports AVACONT_Core
Imports AVACONT_Core.Logger
Imports AVACONT_Core.modTypes

Namespace AVACONT_Viewer
    Partial Public Class MainWindow
        Inherits Window

        Public Sub New()
            InitializeComponent()
            Try
                BuildTreeAsync()
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Eroare la inițializare", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End Sub

        Private Async Sub BuildTreeAsync()
            Try
                treeMain.Items.Clear()

                ' === 1️⃣ Rulează în fundal doar extragerea datelor simple ===
                Dim moduleInfos = Await Task.Run(Function()
                                                     Return GlobalModules.Values.Select(Function(m) New With {
                                                 .Name = m.Name,
                                                 .Type = m.Type,
                                                 .Tag = m
                                             }).ToList()
                                                 End Function)

                Dim formInfos = Await Task.Run(Function()
                                                   Return GlobalFormsReports.Values.Select(Function(f) New With {
                                               .Name = f.Name,
                                               .Type = f.Type,
                                               .Tag = f
                                           }).ToList()
                                               End Function)

                ' === 2️⃣ Construiește controalele WPF DOAR pe UI thread ===
                Await Dispatcher.InvokeAsync(Sub()
                                                 ' === MODULES ===
                                                 For Each m In moduleInfos
                                                     Dim node As New TreeViewItem With {
                                                 .Header = $"📦 {m.Name} ({m.Type})",
                                                 .Tag = m.Tag
                                             }
                                                     node.Items.Add("Loading...")
                                                     AddHandler node.Expanded, AddressOf LazyExpandNode
                                                     treeMain.Items.Add(node)
                                                 Next

                                                 ' === FORMS / REPORTS ===
                                                 For Each f In formInfos
                                                     Dim node As New TreeViewItem With {
                                                 .Header = $"🧩 {f.Name} ({f.Type})",
                                                 .Tag = f.Tag
                                             }
                                                     node.Items.Add("Loading...")
                                                     AddHandler node.Expanded, AddressOf LazyExpandNode
                                                     treeMain.Items.Add(node)
                                                 Next
                                             End Sub)

                LogInfo("✅ BuildTreeAsync finalizat fără erori.")
            Catch ex As Exception
                LogError("BuildTreeAsync", ex)
                MessageBox.Show($"Eroare la construirea arborelui: {ex.Message}", "Eroare", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End Sub

        ' ======================================================================
        ' BUILD TREE (complet populat)
        ' ======================================================================
        Private Sub BuildTree()
            treeMain.Items.Clear()

            ' === MODULES ===
            For Each kvp In GlobalModules
                Dim m = kvp.Value
                Dim node = New TreeViewItem With {.Header = $"📦 {m.Name} ({m.Type})", .Tag = m}

                ' Populate module content
                AddModuleContainerNodes(node, m)
                treeMain.Items.Add(node)
                LogInfo($"{m.Name} adaugat in copac")
            Next

            ' === FORMS / REPORTS ===
            For Each kvp In GlobalFormsReports
                Dim f = kvp.Value
                Dim node = New TreeViewItem With {.Header = $"🧩 {f.Name} ({f.Type})", .Tag = f}

                AddFormReportContainerNodes(node, f)
                treeMain.Items.Add(node)
                LogInfo($"{f.Name} adaugat in copac")

            Next
        End Sub

        ' ======================================================================
        ' MODULE CONTAINER STRUCTURE
        ' ======================================================================
        Private Sub AddModuleContainerNodes(parent As TreeViewItem, m As ModuleContainer)
            AddDeclaresNode(parent, m.Declares)
            AddParamNodes(parent, "📄 Variables", m.Variables)
            AddParamNodes(parent, "📜 Constants", m.Constants)
            AddMethodsNode(parent, m.Methods)
            AddLinesNode(parent, m.Lines)
            AddTokensNode(parent, Nothing)
        End Sub

        ' ======================================================================
        ' FORM / REPORT STRUCTURE
        ' ======================================================================
        Private Sub AddFormReportContainerNodes(parent As TreeViewItem, f As FormReportContainer)
            AddDeclaresNode(parent, f.Declares)
            AddParamNodes(parent, "📄 Variables", f.Variables)
            AddParamNodes(parent, "📜 Constants", f.Constants)
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
        Private Sub AddDeclaresNode(parent As TreeViewItem, declares As Dictionary(Of String, MethodInfo))
            If declares Is Nothing OrElse declares.Count = 0 Then Return
            Dim dNode = New TreeViewItem With {.Header = "⚙️ Declares"}
            For Each kvp In declares
                Dim m = kvp.Value
                Dim mNode = New TreeViewItem With {.Header = $"🔹 {m.Name}", .Tag = m}
                AddTokensNode(mNode, m.TokenizedLine?.Tokens)
                dNode.Items.Add(mNode)
            Next
            parent.Items.Add(dNode)
        End Sub

        ' ======================================================================
        ' VARIABLES / CONSTANTS
        ' ======================================================================
        Private Sub AddParamNodes(parent As TreeViewItem, header As String, list As List(Of ParamInfo))
            If list Is Nothing OrElse list.Count = 0 Then Return
            Dim pNode = New TreeViewItem With {.Header = header}
            For Each p In list
                Dim vNode = New TreeViewItem With {.Header = $"🔹 {p.Name} : {p.Type}", .Tag = p}
                AddTokensNode(vNode, p.Tokens)
                pNode.Items.Add(vNode)
            Next
            parent.Items.Add(pNode)
        End Sub

        ' ======================================================================
        ' METHODS
        ' ======================================================================
        Private Sub AddMethodsNode(parent As TreeViewItem, methods As Dictionary(Of String, MethodInfo))
            If methods Is Nothing OrElse methods.Count = 0 Then Return
            Dim mNode = New TreeViewItem With {.Header = "🧠 Methods"}
            For Each kvp In methods
                Dim method = kvp.Value
                Dim mSub = New TreeViewItem With {.Header = $"🔸 {method.Name}", .Tag = method}

                ' Tokens (header line)
                AddTokensNode(mSub, method.TokenizedLine?.Tokens)

                ' Parameters
                AddParamNodes(mSub, "🧩 Parameters", method.Parameters)

                ' Variables
                AddParamNodes(mSub, "📄 Variables", method.Variables)

                ' Constants
                AddParamNodes(mSub, "📜 Constants", method.Constants)

                ' MethodLines
                If method.MethodLines IsNot Nothing AndAlso method.MethodLines.Count > 0 Then
                    Dim linesNode = New TreeViewItem With {.Header = "📑 MethodLines"}
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

        ' ======================================================================
        ' LINES NODE
        ' ======================================================================
        Private Sub AddLinesNode(parent As TreeViewItem, lines As List(Of MethodLine))
            If lines Is Nothing OrElse lines.Count = 0 Then Return
            Dim node = New TreeViewItem With {.Header = "🪜 Lines"}
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
        Private Sub AddTokensNode(parent As TreeViewItem, tokens As List(Of Token))
            If tokens Is Nothing OrElse tokens.Count = 0 Then Return
            Dim tNode = New TreeViewItem With {.Header = "💬 Tokens"}
            For Each t In tokens
                tNode.Items.Add(New TreeViewItem With {.Header = $"{t.TokenType}: {t.TokenString}", .Tag = t})
            Next
            parent.Items.Add(tNode)
        End Sub

        ' =============================================================
        ' ON CLICK: SHOW DETAILS
        ' =============================================================
        Private Sub TreeMain_SelectedItemChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object))
            lstSubElements.Items.Clear()
            txtSubDetail.Clear()
            txtDetails.Clear()

            Dim node = TryCast(treeMain.SelectedItem, TreeViewItem)
            If node Is Nothing OrElse node.Tag Is Nothing Then Return

            Dim sb As New StringBuilder()

            Select Case True
                Case TypeOf node.Tag Is ModuleContainer
                    Dim m = CType(node.Tag, ModuleContainer)
                    sb.AppendLine($"📦 MODULE CONTAINER DETAILS")
                    sb.AppendLine($"Name: {m.Name}")
                    sb.AppendLine($"Type: {m.Type}")
                    sb.AppendLine($"File Path: {m.FilePath}")
                    sb.AppendLine($"No. of Lines: {If(m.Lines?.Count, 0)}")
                    sb.AppendLine($"No. of Methods: {If(m.Methods?.Count, 0)}")
                    sb.AppendLine($"No. of Constants: {If(m.Constants?.Count, 0)}")
                    sb.AppendLine($"No. of Declares: {If(m.Declares?.Count, 0)}")
                    sb.AppendLine(vbCrLf & "===== CODE =====")
                    sb.AppendLine(m.CodeContent)
                    txtDetails.Text = sb.ToString()

                Case TypeOf node.Tag Is FormReportContainer
                    Dim f = CType(node.Tag, FormReportContainer)
                    sb.AppendLine($"🧩 FORM/REPORT CONTAINER DETAILS")
                    sb.AppendLine($"Name: {f.Name}")
                    sb.AppendLine($"Type: {f.Type}")
                    sb.AppendLine($"File Path: {f.FilePath}")
                    sb.AppendLine($"Sections: {f.Sections.Count}")
                    sb.AppendLine($"Controls: {f.Controls.Count}")
                    sb.AppendLine($"Methods: {If(f.Methods?.Count, 0)}")
                    sb.AppendLine(vbCrLf & "===== CODE =====")
                    sb.AppendLine(f.CodeContent)
                    txtDetails.Text = sb.ToString()

                Case TypeOf node.Tag Is MethodInfo
                    Dim m = CType(node.Tag, MethodInfo)
                    sb.AppendLine($"🧠 METHOD DETAILS")
                    sb.AppendLine($"Name: {m.Name}")
                    sb.AppendLine($"Signature: {m.Signature}")
                    sb.AppendLine($"Scope: {m.MethodScope}")
                    sb.AppendLine($"StartLine: {m.StartLine} | EndLine: {m.EndLine}")
                    sb.AppendLine($"Parameters: {If(m.Parameters?.Count, 0)}")
                    sb.AppendLine($"Variables: {If(m.Variables?.Count, 0)}")
                    sb.AppendLine($"Constants: {If(m.Constants?.Count, 0)}")
                    sb.AppendLine(vbCrLf & "===== CODE =====")
                    sb.AppendLine(m.CodeContent)
                    txtDetails.Text = sb.ToString()

                    ' === Populate ListBox with MethodLines & Tokens ===
                    If m.MethodLines IsNot Nothing Then
                        For Each l In m.MethodLines
                            lstSubElements.Items.Add($"Line {l.LineNumber}: {l.Content}")
                        Next
                    End If
                    If m.Tokens IsNot Nothing Then
                        For Each t In m.Tokens
                            lstSubElements.Items.Add($"Token: {t.TokenString} ({t.TokenType})")
                        Next
                    End If

                Case TypeOf node.Tag Is ParamInfo
                    Dim v = CType(node.Tag, ParamInfo)
                    sb.AppendLine($"🔹 VARIABLE / PARAMETER")
                    sb.AppendLine($"Name: {v.Name}")
                    sb.AppendLine($"Type: {v.Type}")
                    sb.AppendLine($"Value: {v.Value}")
                    txtDetails.Text = sb.ToString()

                    ' Tokens (dacă există)
                    If v.Tokens IsNot Nothing Then
                        For Each t In v.Tokens
                            lstSubElements.Items.Add($"{t.TokenType}: {t.TokenString}")
                        Next
                    End If

                Case TypeOf node.Tag Is MethodLine
                    Dim ml = CType(node.Tag, MethodLine)
                    sb.AppendLine($"📄 METHOD LINE")
                    sb.AppendLine($"Line: {ml.LineNumber}")
                    sb.AppendLine(ml.Content)
                    txtDetails.Text = sb.ToString()
                    If ml.Tokens IsNot Nothing Then
                        For Each t In ml.Tokens
                            lstSubElements.Items.Add($"{t.TokenType}: {t.TokenString}")
                        Next
                    End If

                Case TypeOf node.Tag Is Token
                    Dim t = CType(node.Tag, Token)
                    sb.AppendLine($"💬 TOKEN DETAIL")
                    sb.AppendLine($"Token: {t.TokenString}")
                    sb.AppendLine($"Type: {t.TokenType}")
                    sb.AppendLine($"Context: {t.Context}")
                    sb.AppendLine($"Line: {t.LineNumber}")
                    txtDetails.Text = sb.ToString()

                Case TypeOf node.Tag Is FormControl
                    Dim c = CType(node.Tag, FormControl)
                    sb.AppendLine($"🎛 CONTROL")
                    sb.AppendLine($"Name: {c.Name}")
                    sb.AppendLine($"Type: {c.Type}")
                    sb.AppendLine($"Section: {If(c.Section?.Name, "-")}")
                    sb.AppendLine($"Caption: {c.Caption}")
                    sb.AppendLine($"ControlSource: {c.ControlSource}")
                    sb.AppendLine($"RecordSource: {c.RecordSource}")
                    txtDetails.Text = sb.ToString()

                Case Else
                    txtDetails.Text = $"ℹ️ {node.Header}"
            End Select
        End Sub

        ' =============================================================
        ' SUBELEMENT SELECTED (from ListBox)
        ' =============================================================
        Private Sub LstSubElements_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            Dim sel = TryCast(lstSubElements.SelectedItem, String)
            If String.IsNullOrEmpty(sel) Then Return

            ' Simple text expansion for now (could map to Tag objects if needed)
            txtSubDetail.Text = $"🔍 Selected: {sel}" & vbCrLf &
                        "–––––––––––––––––––––––" & vbCrLf &
                        "Aici poți adăuga mai târziu detalii specifice (token info, context, tip etc.)"
        End Sub

        ' În MainWindow.xaml.vb, adaugă aceste metode noi:

        Private Async Sub LazyExpandNode(sender As Object, e As RoutedEventArgs)
            Dim node = TryCast(sender, TreeViewItem)
            If node Is Nothing OrElse node.Tag Is Nothing Then Return

            ' ➜ continuăm DOAR dacă nodul e în stare de "loading"
            If Not IsNodeLoading(node) Then Return

            ' snapshot sigur
            Dim tagObject As Object = node.Tag

            ' arată spinner
            ShowLoadingState(node)

            Dim loadEx As Exception = Nothing
            Dim childData As List(Of (Header As String, Tag As Object, Loader As Action(Of TreeViewItem))) = Nothing

            Try
                ' === date în fundal ===
                childData = Await Task.Run(Function() CollectChildData(tagObject))
            Catch ex As Exception
                loadEx = ex
            End Try

            ' === UI update ===
            Await Dispatcher.InvokeAsync(Sub() PopulateNode(node, tagObject, childData, loadEx))
        End Sub

        Private Function IsNodeLoading(node As TreeViewItem) As Boolean
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

        Private Sub ShowLoadingState(node As TreeViewItem)
            node.Items.Clear()
            node.Items.Add(New TreeViewItem With {.Header = "⏳ Se încarcă..."})
            node.IsExpanded = True
        End Sub

        Private Function CollectChildData(tagObject As Object) As List(Of (Header As String, Tag As Object, Loader As Action(Of TreeViewItem)))
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

        Private Sub PopulateNode(node As TreeViewItem, tagObject As Object, childData As List(Of (Header As String, Tag As Object, Loader As Action(Of TreeViewItem))), loadEx As Exception)
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
                PopulateModuleNode(node, CType(tagObject, ModuleContainer))
            ElseIf TypeOf tagObject Is FormReportContainer Then
                PopulateFormReportNode(node, CType(tagObject, FormReportContainer))
            End If
        End Sub

        Private Sub PopulateModuleNode(node As TreeViewItem, m As ModuleContainer)
            If m.Declares?.Count > 0 Then AddDeclaresNode(node, m.Declares)
            If m.Variables?.Count > 0 Then AddParamNodes(node, "📄 Variables", m.Variables)
            If m.Constants?.Count > 0 Then AddParamNodes(node, "📜 Constants", m.Constants)
            If m.Methods?.Count > 0 Then AddMethodsNode(node, m.Methods)
            If m.Lines?.Count > 0 Then AddLinesNode(node, m.Lines)
        End Sub

        Private Sub PopulateFormReportNode(node As TreeViewItem, f As FormReportContainer)
            If f.Declares?.Count > 0 Then AddDeclaresNode(node, f.Declares)
            If f.Variables?.Count > 0 Then AddParamNodes(node, "📄 Variables", f.Variables)
            If f.Constants?.Count > 0 Then AddParamNodes(node, "📜 Constants", f.Constants)
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


        Private Sub TxtSearch_TextChanged(sender As Object, e As TextChangedEventArgs)
            Dim query = txtSearch.Text.Trim()
            If query.Length < 2 Then Return

            ' caută în paralel în tot proiectul
            Dim match = GlobalModules.Values.SelectMany(Function(m) m.Methods.Values.Select(Function(meth) (m, meth))) _
            .FirstOrDefault(Function(x) x.meth.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)

            If match.m IsNot Nothing Then
                HighlightTreeNode(match.m.Name, match.meth.Name)
            End If
        End Sub

        Private Sub HighlightTreeNode(moduleName As String, methodName As String)
            For Each n As TreeViewItem In treeMain.Items
                Dim m = TryCast(n.Tag, ModuleContainer)
                Dim f = TryCast(n.Tag, FormReportContainer)

                Dim name = If(m?.Name, f?.Name)
                If String.Equals(name, moduleName, StringComparison.OrdinalIgnoreCase) Then
                    ' expandează nodul modulului
                    n.IsExpanded = True
                    If n.Items.Count > 0 Then LazyExpandNode(n, Nothing)

                    ' caută metoda în ramura Methods
                    For Each subNode As TreeViewItem In n.Items
                        If subNode.Header.ToString().Contains("Methods") Then
                            If subNode.Items.Count = 0 Then LazyExpandNode(subNode, Nothing)
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

    End Class
End Namespace
