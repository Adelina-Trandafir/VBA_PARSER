'PROJECT NAME: VBA_LAUNCHER
'FILE DESCRIPTION: Vizualizarea principală pentru analiza și vizualizarea conținutului VBA din fișiere Access
'PATH: VBA_LAUNCHER/MainWindow.xaml.vb

Imports Microsoft.Win32
Imports VBA_CORE
Imports VBA_CORE.modGlobals

Namespace VBA_LAUNCHER
    Class MainWindow
        Private isParsed As Boolean = False
        Private isFiltering As Boolean = False

        Private viewerWindow As VBA_VIEWER.VBA_VIEWER.MainWindow = Nothing
        Private ReadOnly allLogLines As New List(Of String)()

        Public Sub New()
            InitializeComponent()
            WireViewerWindowClosed()
            AddHandler Logger.OnLogMessage, AddressOf HandleLogMessage
            Logger.UI_InvokeSafe = Sub(msg As String, lvl As Integer)
                                       Dispatcher.Invoke(
                                           Sub()
                                               HandleLogMessage(msg, lvl)
                                           End Sub)
                                   End Sub
            txtLog.Document.PagePadding = New Thickness(0)
        End Sub

        Private ReadOnly Property TailParagraph As Paragraph
            Get
                Dim p = TryCast(txtLog.Document.Blocks.LastBlock, Paragraph)
                If p Is Nothing Then
                    p = New Paragraph() With {.Margin = New Thickness(0), .LineHeight = 14}
                    txtLog.Document.Blocks.Add(p)
                End If
                Return p
            End Get
        End Property

        Private Sub WireViewerWindowClosed()
            If viewerWindow IsNot Nothing Then
                RemoveHandler viewerWindow.Closed, AddressOf ViewerWindow_Closed
            End If
            If viewerWindow IsNot Nothing Then
                AddHandler viewerWindow.Closed, AddressOf ViewerWindow_Closed
            End If

        End Sub
        Private Sub ViewerWindow_Closed(sender As Object, e As EventArgs)
            viewerWindow = Nothing
        End Sub

        Private Sub Browse_Click(sender As Object, e As RoutedEventArgs)
            Dim dlg As New OpenFileDialog With {.Filter = "Access DB|*.accdb;*.mdb"}
            If dlg.ShowDialog() Then txtPath.Text = dlg.FileName
        End Sub

        Private Async Sub Start_Click(sender As Object, e As RoutedEventArgs)
            ExtendedVerbose = chkExtendedVerbose.IsChecked.GetValueOrDefault()

            btnStart.IsEnabled = False
            Me.btnFwd.Visibility = Visibility.Hidden
            txtSearch.Visibility = Visibility.Collapsed
            txtSearch.Text = String.Empty
            allLogLines.Clear()
            isFiltering = False

            lblStatus.Text = "🔄 Rulez analiza..."

            modGlobals.Global_DBPath = txtPath.Text
            modGlobals.Global_UseCache = chkCache.IsChecked
            modGlobals.Global_UseParallel = chkParallel.IsChecked
            modGlobals.Global_VerboseLevel = If(rbQuiet.IsChecked, 0, If(rbVerbose.IsChecked, 2, 1))

            Await Task.Run(Sub() VBA_TOKENIZER.Program.RunFullAnalysis())

            lblStatus.Text = "✅ Analiză completă!"
            isParsed = True
            btnStart.IsEnabled = True
            txtSearch.Visibility = Visibility.Visible

            ' ✅ recreează dacă a fost închisă
            If viewerWindow Is Nothing OrElse Not viewerWindow.IsLoaded Then
                viewerWindow = New VBA_VIEWER.VBA_VIEWER.MainWindow With {.Tag = Me, .Owner = Me}
                AddHandler viewerWindow.Closed, AddressOf ViewerWindow_Closed
            End If

            ' arată sau readuce în față
            If viewerWindow.IsVisible Then
                viewerWindow.Activate()
            Else
                viewerWindow.Show()
            End If

            Me.btnFwd.Visibility = Visibility.Visible
            Me.Hide()
        End Sub

        Private Sub HandleLogMessage(msg As String, level As Integer)
            If modGlobals.Global_VerboseLevel = 0 Then Exit Sub

            Dim isParallel As Boolean = chkParallel.IsChecked.GetValueOrDefault(False)
            Dim indentLevel As Integer = Math.Max(0, level - 1)

            ' 🔹 Definește explicit delegatul Action pentru Invoke
            Dim action As Action =
                            Sub()
                                Dim L As String = msg.ToLowerInvariant()
                                Dim isCounter As Boolean =
                                    (Not isParallel) AndAlso
                                    (L.Contains("(module)") OrElse L.Contains("(form)") OrElse L.Contains("(report)") OrElse L.Contains("(class)"))

                                If isCounter Then
                                    AppendObjectCounter(msg, indentLevel)
                                Else
                                    AddLine(msg, Nothing, False, indentLevel)
                                End If

                                txtLog.ScrollToEnd()
                            End Sub

            ' 🔹 Apelează Invoke cu overloadul corect
            Dispatcher.Invoke(action, System.Windows.Threading.DispatcherPriority.Background)
        End Sub

        Private Sub btnFwd_Click(sender As Object, e As RoutedEventArgs) Handles btnFwd.Click
            ' 🔹 Lansează Viewer-ul, dar ascunde fereastra curentă (nu o închide)
            If viewerWindow Is Nothing Then viewerWindow = New VBA_VIEWER.VBA_VIEWER.MainWindow With {.Tag = Me}
            viewerWindow.Show()
            Me.Hide()

        End Sub

        Private Sub chkCache_Click(sender As Object, e As RoutedEventArgs) Handles chkCache.Click
            chkParallel.IsEnabled = Not chkCache.IsChecked.GetValueOrDefault()
            LogOptions.IsEnabled = Not chkCache.IsChecked.GetValueOrDefault()
            txtPath.IsEnabled = Not chkCache.IsChecked.GetValueOrDefault()
            btnBrowse.IsEnabled = Not chkCache.IsChecked.GetValueOrDefault()

            'If chkParallel.IsChecked Then
            '    RemoveHandler Logger.OnLogMessage, AddressOf HandleLogMessage
            'Else
            '    AddHandler Logger.OnLogMessage, AddressOf HandleLogMessage
            'End If
        End Sub

        Private Sub txtSearch_TextChanged(sender As Object, e As TextChangedEventArgs)
            RebuildLog(txtSearch.Text)

        End Sub

        Private Sub Verbose_Checked(sender As Object, e As RoutedEventArgs)
            chkExtendedVerbose.Visibility = Visibility.Visible
            chkExtendedVerbose.IsEnabled = True
        End Sub

        Private Sub Verbose_Unchecked(sender As Object, e As RoutedEventArgs)
            chkExtendedVerbose.Visibility = Visibility.Hidden
            chkExtendedVerbose.IsChecked = False
            chkExtendedVerbose.IsEnabled = False
        End Sub

        '==================================================================
        '==== HELPER FUNCTIONS ===========================================
        '==================================================================
        Private Sub RebuildLog(filter As String)
            Dim f As String = If(filter, String.Empty).Trim()
            isFiltering = f.Length > 0

            ' === reconstruim documentul ===
            Dim doc As New FlowDocument() With {
                .PagePadding = New Thickness(0),
                .FontFamily = New FontFamily("Consolas"),
                .FontSize = 12,
                .LineHeight = 13,
                .LineStackingStrategy = LineStackingStrategy.BlockLineHeight
            }

            ' ✅ corect: adăugăm Setter prin Add()
            Dim style As New Style(GetType(Paragraph))
            style.Setters.Add(New Setter(Paragraph.MarginProperty, New Thickness(0)))
            doc.Resources.Add(GetType(Paragraph), style)

            Dim p As New Paragraph()
            doc.Blocks.Add(p)

            Dim q = If(isFiltering,
               allLogLines.Where(Function(s) s.IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0),
               allLogLines.AsEnumerable())

            For Each line In q
                Dim fore As Brush = Brushes.Black
                Dim L = line.ToLowerInvariant()
                If L.Contains("(form)") Then fore = Brushes.MediumSeaGreen
                If L.Contains("(report)") Then fore = Brushes.OrangeRed
                If L.Contains("(class)") Then fore = Brushes.SteelBlue
                If L.Contains("(module)") Then fore = Brushes.SlateGray

                p.Inlines.Add(New Run(line) With {.Foreground = fore})
                p.Inlines.Add(New LineBreak())
            Next

            txtLog.Document = doc
            txtLog.ScrollToEnd()
        End Sub

        Private Sub AddLine(msg As String, Optional fore As Brush = Nothing, Optional bold As Boolean = False, Optional indentLevel As Integer = 0)
            Dim p = TailParagraph

            ' indent vizual
            If indentLevel > 1 Then
                p.Inlines.Add(New Run(New String(ControlChars.Tab, indentLevel)))
            End If

            Dim r As New Run(msg)
            If fore IsNot Nothing Then r.Foreground = fore
            If bold Then r.FontWeight = FontWeights.Bold

            If chkCache.IsChecked Then
                p.Inlines.Clear()
                p.Inlines.Add(r)
            Else
                p.Inlines.Add(r)
                p.Inlines.Add(New LineBreak())
                ' poți alege să salvezi în cache cu sau fără indent;
                ' ca să păstrăm afișarea identică și la rebuild, salvăm cu indent:
                'allLogLines.Add(New String(ControlChars.Tab, indentLevel) & msg)
                allLogLines.Add(msg)
                If isFiltering Then RebuildLog(txtSearch.Text)
            End If
        End Sub

        Private Sub AppendObjectCounter(counterText As String, Optional indentLevel As Integer = 0)
            Dim parts = counterText.Split(";"c)
            If parts.Length = 0 Then Return

            Dim p = TailParagraph

            ' indent vizual (NU modificăm textul pentru parsare/colorare)
            If indentLevel > 0 Then
                p.Inlines.Add(New Run(New String(ControlChars.Tab, indentLevel)))
            End If

            Dim headerText = parts(0).Trim()

            ' colors (neschimbat)
            Dim headerColor As Brush = Brushes.DarkSlateBlue
            Select Case True
                Case headerText.ToLower().Contains("(form)") : headerColor = Brushes.MediumSeaGreen
                Case headerText.ToLower().Contains("(report)") : headerColor = Brushes.OrangeRed
                Case headerText.ToLower().Contains("(class)") : headerColor = Brushes.SteelBlue
                Case headerText.ToLower().Contains("(module)") : headerColor = Brushes.SlateGray
            End Select

            p.Inlines.Add(New Run(headerText) With {.FontWeight = FontWeights.Bold, .Foreground = headerColor})

            ' flat cache line (inclusiv indent ca să păstrăm afișarea la rebuild)
            Dim flat As New System.Text.StringBuilder(New String(ControlChars.Tab, indentLevel) & headerText)

            For i = 1 To parts.Length - 1
                Dim segment = parts(i).Trim()
                If String.IsNullOrWhiteSpace(segment) Then Continue For

                Dim colonPos = segment.IndexOf(":"c)
                If colonPos > 0 Then
                    Dim label = segment.Substring(0, colonPos).Trim()
                    Dim value = segment.Substring(colonPos + 1).Trim()
                    If value = "0" Then Continue For

                    Dim valueColor As Brush = Brushes.Gray
                    Select Case label.ToLowerInvariant()
                        Case "methods" : valueColor = Brushes.DarkCyan
                        Case "declares" : valueColor = Brushes.SteelBlue
                        Case "variables" : valueColor = Brushes.DarkOliveGreen
                        Case "constants" : valueColor = Brushes.SaddleBrown
                        Case "total tokens" : valueColor = Brushes.DarkMagenta
                        Case "sections" : valueColor = Brushes.Bisque
                        Case "controls" : valueColor = Brushes.GreenYellow
                    End Select

                    p.Inlines.Add(New Run("  " & label & ": ") With {.Foreground = Brushes.Gray})
                    p.Inlines.Add(New Run(value) With {.FontWeight = FontWeights.Bold, .Foreground = valueColor})

                    flat.Append("  ").Append(label).Append(": ").Append(value)
                Else
                    p.Inlines.Add(New Run("  " & segment) With {.Foreground = Brushes.Gray})
                    flat.Append("  ").Append(segment)
                End If
            Next

            p.Inlines.Add(New LineBreak())

            allLogLines.Add(flat.ToString())
            If isFiltering Then RebuildLog(txtSearch.Text)
        End Sub

        Protected Overrides Sub OnClosing(e As ComponentModel.CancelEventArgs)
            MyBase.OnClosing(e)
            Logger.ShutdownLogger()
        End Sub
    End Class
End Namespace
