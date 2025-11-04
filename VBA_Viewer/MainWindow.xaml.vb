'PROJECT NAME: VBA_VIEWER
'FILE DESCRIPTION: Fereastra principală a launcher-ului VBA Analyzer
'PATH: VBA_LAUNCHER/MainWindow.xaml.vb
Imports System.ComponentModel
Imports System.Text
Imports VBA_CORE.Logger
Imports VBA_CORE.modTypes
Imports VBA_MethodInfo = VBA_CORE.modTypes.MethodInfo
Imports VBA_PropertyInfo = VBA_CORE.modTypes.PropertyInfo
Imports VBA_FunctionInfo = VBA_CORE.modTypes.FunctionInfo

Namespace VBA_VIEWER
    Partial Public Class MainWindow
        Inherits Window

        Public moduleInfos = Nothing
        Public formInfos = Nothing
        Public TokenReferences As New Dictionary(Of String, Token)()

        Public Sub New()
            InitializeComponent()
            Try
                BuildTreeAsync()
            Catch ex As Exception
                MessageBox.Show(ex.Message, "Eroare la inițializare", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End Sub
        ' MainWindow.xaml.vb
        Private Async Sub BuildTreeAsync()
            Try
                Await TreeHelpers.BuildTreeAsync(Me)
                LogInfo("✅ BuildTreeAsync completed.")
            Catch ex As Exception
                LogError("BuildTreeAsync", ex)
                MessageBox.Show($"Error building the treecontrol: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error)
            End Try
        End Sub

        Public Async Sub LazyExpandNode(sender As Object, e As RoutedEventArgs)
            Dim node = TryCast(sender, TreeViewItem)
            If node Is Nothing OrElse node.Tag Is Nothing Then Return

            ' ➜ continuăm DOAR dacă nodul e în stare de "loading"
            If Not TreeHelpers.IsNodeLoading(node) Then Return

            ' snapshot sigur
            Dim tagObject As Object = node.Tag

            ' arată spinner
            TreeHelpers.ShowLoadingState(node)

            Dim loadEx As Exception = Nothing
            Dim childData As List(Of (Header As String, Tag As Object, Loader As Action(Of TreeViewItem))) = Nothing

            Try
                ' === date în fundal ===
                childData = Await Task.Run(Function() TreeHelpers.CollectChildData(tagObject))
            Catch ex As Exception
                loadEx = ex
            End Try

            ' === UI update ===
            Await Dispatcher.InvokeAsync(Sub() TreeHelpers.PopulateNode(node, tagObject, childData, loadEx))
        End Sub

        Private Sub TreeMain_SelectedItemChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object))
            lstSubElements.Items.Clear()
            txtSubDetail.Document.Blocks.Clear()
            txtDetails.Clear()

            Dim node = TryCast(treeMain.SelectedItem, TreeViewItem)
            If node Is Nothing OrElse node.Tag Is Nothing Then Return

            Select Case True

        ' =============================================================
        ' MODULE CONTAINER
        ' =============================================================
                Case TypeOf node.Tag Is ModuleContainer
                    Dim m = CType(node.Tag, ModuleContainer)
                    Dim sb As New StringBuilder()
                    Dim countTokens As Integer = 0

                    txtDetails.AppendText(m.CodeContent)
                    sb.AppendLine($"~📦 MODULE CONTAINER DETAILS")
                    sb.AppendLine($"~~Name: {m.Name}")
                    sb.AppendLine($"~~XType: {m.Type}")
                    sb.AppendLine($"-File Path: {m.FilePath}")
                    sb.AppendLine($"=Lines: {If(m.Lines?.Count, 0)}")
                    sb.AppendLine($"=Methods: {If(m.Methods?.Count, 0)}")
                    sb.AppendLine($"=Functions: {If(m.Functions?.Count, 0)}")
                    sb.AppendLine($"=Properties: {If(m.Properties?.Count, 0)}")
                    sb.AppendLine($"=Constants: {If(m.Constants?.Count, 0)}")
                    sb.AppendLine($"=Declares: {If(m.Declares?.Count, 0)}")
                    For Each l In m.Lines
                        countTokens += l.Tokens.Count
                    Next
                    If countTokens > 0 Then sb.AppendLine($"=Tokens:{countTokens}")

                    RichTextHelpers.WriteLinesToRTFDocument(Me, sb)

        ' =============================================================
        ' FORM / REPORT CONTAINER
        ' =============================================================
                Case TypeOf node.Tag Is FormReportContainer
                    Dim countTokens As Integer = 0
                    Dim f = CType(node.Tag, FormReportContainer)
                    txtDetails.AppendText(f.CodeContent)

                    Dim sb As New StringBuilder()
                    sb.AppendLine($"~🧩 FORM/REPORT CONTAINER DETAILS")
                    sb.AppendLine($"~~Name: {f.Name}")
                    sb.AppendLine($"~~Type: {f.Type}")
                    sb.AppendLine($"-File Path: {f.FilePath}")
                    sb.AppendLine($"=Sections: {f.Sections.Count}")
                    sb.AppendLine($"=Controls: {f.Controls.Count}")
                    sb.AppendLine($"=Methods: {If(f.Methods?.Count, 0)}")
                    sb.AppendLine($"=Functions: {If(f.Functions?.Count, 0)}")
                    sb.AppendLine($"=Properties: {If(f.Properties?.Count, 0)}")
                    For Each l In f.Lines
                        countTokens += l.Tokens.Count
                    Next
                    If countTokens > 0 Then sb.AppendLine($"=Tokens:{countTokens}")

                    RichTextHelpers.WriteLinesToRTFDocument(Me, sb)

        ' =============================================================
        ' METHOD
        ' =============================================================
                Case TypeOf node.Tag Is VBA_MethodInfo
                    Dim m = CType(node.Tag, VBA_MethodInfo)
                    Dim countTokens As Integer = 0
                    txtDetails.AppendText(m.CodeContent)

                    Dim sb As New StringBuilder()
                    sb.AppendLine($"~🧠 METHOD DETAILS")
                    sb.AppendLine($"~~Name: {m.Name}")
                    sb.AppendLine($"~~Scope: {m.MethodScope}")
                    sb.AppendLine($"-Signature: {m.Signature}")
                    sb.AppendLine($"=StartLine: {m.StartLine}")
                    sb.AppendLine($"=EndLine: {m.EndLine}")
                    sb.AppendLine($"=Parameters: {If(m.Parameters?.Count, 0)}")
                    sb.AppendLine($"=Variables: {If(m.Variables?.Count, 0)}")
                    sb.AppendLine($"=Constants: {If(m.Constants?.Count, 0)}")

                    RichTextHelpers.WriteLinesToRTFDocument(Me, sb)

                    If m.MethodLines IsNot Nothing Then
                        For Each l In m.MethodLines
                            countTokens += l.Tokens.Count
                            lstSubElements.Items.Add(New ListBoxItem With {.Content = $"Line {l.LineNumber}: {l.Content}", .Tag = l})
                        Next
                        If countTokens > 0 Then sb.AppendLine($"=Tokens:{countTokens}")
                    End If

        ' =============================================================
        ' FUNCTION
        ' =============================================================
                Case TypeOf node.Tag Is VBA_FunctionInfo
                    Dim f = CType(node.Tag, VBA_FunctionInfo)
                    txtDetails.AppendText(f.CodeContent)

                    Dim sb As New StringBuilder()
                    sb.AppendLine($"~🔧 FUNCTION DETAILS")
                    sb.AppendLine($"~~Name: {f.Name}")
                    sb.AppendLine($"~~Scope: {f.MethodScope}")
                    sb.AppendLine($"~~Return Type: {f.ReturnType}")
                    sb.AppendLine($"-Signature: {f.Signature}")
                    sb.AppendLine($"=StartLine: {f.StartLine}")
                    sb.AppendLine($"=EndLine: {f.EndLine}")
                    sb.AppendLine($"=Parameters: {If(f.Parameters?.Count, 0)}")
                    sb.AppendLine($"=Variables: {If(f.Variables?.Count, 0)}")
                    sb.AppendLine($"=Constants: {If(f.Constants?.Count, 0)}")

                    RichTextHelpers.WriteLinesToRTFDocument(Me, sb)

                    If f.MethodLines IsNot Nothing Then
                        For Each l In f.MethodLines
                            lstSubElements.Items.Add(New ListBoxItem With {.Content = $"Line {l.LineNumber}: {l.Content}", .Tag = l})
                        Next
                    End If

        ' =============================================================
        ' PROPERTY
        ' =============================================================
                Case TypeOf node.Tag Is VBA_PropertyInfo
                    Dim p = CType(node.Tag, VBA_PropertyInfo)
                    txtDetails.AppendText(p.CodeContent)

                    Dim sb As New StringBuilder()
                    sb.AppendLine($"~🧩 PROPERTY DETAILS")
                    sb.AppendLine($"~~Name: {p.Name}")
                    sb.AppendLine($"~~Scope: {p.MethodScope}")
                    sb.AppendLine($"~~Property Type: {p.PropertyType}")
                    sb.AppendLine($"~~Return Type: {p.ReturnType}")
                    sb.AppendLine($"-Signature: {p.Signature}")
                    sb.AppendLine($"=StartLine: {p.StartLine}")
                    sb.AppendLine($"=EndLine: {p.EndLine}")
                    sb.AppendLine($"=Parameters: {If(p.Parameters?.Count, 0)}")
                    sb.AppendLine($"=Variables: {If(p.Variables?.Count, 0)}")
                    sb.AppendLine($"=Constants: {If(p.Constants?.Count, 0)}")

                    RichTextHelpers.WriteLinesToRTFDocument(Me, sb)

                    If p.MethodLines IsNot Nothing Then
                        For Each l In p.MethodLines
                            lstSubElements.Items.Add(New ListBoxItem With {.Content = $"Line {l.LineNumber}: {l.Content}", .Tag = l})
                        Next
                    End If

        ' =============================================================
        ' PARAMLIST / LINES / TOKENS
        ' =============================================================
                Case TryCast(node.Tag, List(Of ParamInfo)) IsNot Nothing
                    Dim paramList = TryCast(node.Tag, List(Of ParamInfo))
                    For Each p In paramList
                        lstSubElements.Items.Add(New ListBoxItem With {.Content = $"🔹 {p.Name} : {p.Type}", .Tag = p})
                    Next

                Case TryCast(node.Tag, List(Of MethodLine)) IsNot Nothing
                    Dim lines = TryCast(node.Tag, List(Of MethodLine))
                    For Each line In lines
                        lstSubElements.Items.Add(New ListBoxItem With {.Content = $"🔹 {line.LineNumber} : {line.Content}", .Tag = line})
                    Next

                Case TryCast(node.Tag, List(Of Token)) IsNot Nothing
                    Dim tokens = TryCast(node.Tag, List(Of Token))
                    For Each token In tokens
                        lstSubElements.Items.Add(New ListBoxItem With {.Content = $"🔹 {token.LineNumber} : {String.Join(":", token.TokenString, token.NextSymbol, token.TokenType)}", .Tag = token})
                    Next
            End Select
        End Sub


        ' =============================================================
        ' SUBELEMENT SELECTED (from ListBox)
        ' =============================================================
        Private Sub LstSubElements_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
            txtSubDetail.Document.Blocks.Clear()

            Dim item = TryCast(lstSubElements.SelectedItem, ListBoxItem)
            If item Is Nothing OrElse item.Tag Is Nothing Then Return

            Dim sb As New StringBuilder()

            Select Case True
                Case TypeOf item.Tag Is MethodLine
                    Dim ml = CType(item.Tag, MethodLine)

                    txtSubDetail.Document = RichTextHelpers.BuildTokenGridDocument(Me, ml)

                Case TypeOf item.Tag Is Token
                    Dim t = CType(item.Tag, Token)
                    Dim doc As New FlowDocument()
                    doc.Blocks.Add(New Paragraph(New Run("💬 TOKEN") With {.FontWeight = FontWeights.Bold}))
                    doc.Blocks.Add(New Paragraph(New Run($"Token: {t.TokenString}")))
                    doc.Blocks.Add(New Paragraph(New Run($"Type: {t.TokenType}")))
                    doc.Blocks.Add(New Paragraph(New Run($"Line: {t.LineNumber}")))
                    doc.Blocks.Add(New Paragraph(New Run($"Context: {t.Context}")))
                    txtSubDetail.Document = doc

                Case TypeOf item.Tag Is ParamInfo
                    Dim p = CType(item.Tag, ParamInfo)
                    Dim doc As New FlowDocument()
                    doc.Blocks.Add(New Paragraph(New Run("🔹 PARAMETER") With {.FontWeight = FontWeights.Bold}))
                    doc.Blocks.Add(New Paragraph(New Run($"Name: {p.Name}")))
                    doc.Blocks.Add(New Paragraph(New Run($"Type: {p.Type}")))
                    doc.Blocks.Add(New Paragraph(New Run($"Value: {p.Value}")))
                    txtSubDetail.Document = doc

                Case Else
                    Dim doc As New FlowDocument()
                    doc.Blocks.Add(New Paragraph(New Run($"ℹ️ Selected: {item.Content}")))
                    txtSubDetail.Document = doc
            End Select

        End Sub
        Public Sub TokenType_Click(sender As Object, e As RoutedEventArgs)
            If TryCast(sender, Hyperlink) IsNot Nothing Then
                Dim link = TryCast(sender, Hyperlink)
                Dim linkTag As String = ""

                If link?.Tag IsNot Nothing Then
                    linkTag = link.Tag

                    Select Case linkTag
                        Case ""
                    End Select
                End If

            End If
        End Sub

        Private Sub TxtSearch_TextChanged(sender As Object, e As TextChangedEventArgs)
            Dim query = txtSearch.Text.Trim()
            If query.Length < 2 Then Return

            ' caută în paralel în tot proiectul
            Dim match = GlobalModules.Values.SelectMany(Function(m) m.Methods.Values.Select(Function(meth) (m, meth))) _
            .FirstOrDefault(Function(x) x.meth.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)

            If match.m IsNot Nothing Then
                TreeHelpers.HighlightTreeNode(Me, treeMain, match.m.Name, match.meth.Name)
            End If
        End Sub

        Private Sub BtnBack_Click(sender As Object, e As RoutedEventArgs)
            Try
                ' Caută fereastra de launcher stocată în Tag
                Dim launcher = TryCast(Me.Tag, Window)
                launcher?.Show()

                ' Închide viewerul
                Me.Hide()

            Catch ex As Exception
                MessageBox.Show("Eroare la revenirea la opțiuni: " & ex.Message,
                                "Eroare",
                                MessageBoxButton.OK,
                                MessageBoxImage.Error)
            End Try
        End Sub

        Private Sub MainWindow_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
            Dim launcher = TryCast(Me.Tag, Window)
            launcher?.Show()

        End Sub

        Private Sub RbNoGroup_Checked(sender As Object, e As RoutedEventArgs)
            If Not IsLoaded Then Return
            BuildTreeAsync()
        End Sub

        Private Sub RbGroupByType_Checked(sender As Object, e As RoutedEventArgs)
            If Not IsLoaded Then Return
            TreeHelpers.BuildTreeGroupedByType(Me, treeMain)
        End Sub

        Private Sub RbGroupByRef_Checked(sender As Object, e As RoutedEventArgs)
            If Not IsLoaded Then Return
            BuildTreeGroupedByReference()
        End Sub

        Private Sub BuildTreeGroupedByReference()
            'TODO: Implementare grupare după referințe
        End Sub

        ' ⭐ HANDLER NOU PENTRU CLICK PE NAME
        ' ⭐ HANDLER PENTRU CLICK PE NAME
        Friend Sub TokenName_Click(sender As Object, e As RoutedEventArgs)
            Dim link = TryCast(sender, Hyperlink)
            If link Is Nothing OrElse link.Tag Is Nothing Then Return

            ' ⭐ Recuperează token-ul din dicționar folosind ID-ul
            Dim tokenId As String = link.Tag.ToString()
            If Not TokenReferences.ContainsKey(tokenId) Then Return

            Dim token As Token = TokenReferences(tokenId)

            ' Afișează popup
            'ShowRTBTokenPropertiesPopup(token)
            ShowTreeTokenPropertiesPopup(token, $"Token: {token.TokenString}")
        End Sub

    End Class
End Namespace