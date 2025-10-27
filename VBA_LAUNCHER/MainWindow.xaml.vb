Imports Microsoft.Win32
Imports VBA_CORE
Imports VBA_CORE.VBA_CORE

Namespace VBA_LAUNCHER
    Class MainWindow
        Public Sub New()
            InitializeComponent()
            AddHandler Logger.OnLogMessage, AddressOf HandleLogMessage
        End Sub

        Private Sub Browse_Click(sender As Object, e As RoutedEventArgs)
            Dim dlg As New OpenFileDialog With {.Filter = "Access DB|*.accdb;*.mdb"}
            If dlg.ShowDialog() Then txtPath.Text = dlg.FileName
        End Sub

        Private Async Sub Start_Click(sender As Object, e As RoutedEventArgs)
            btnStart.IsEnabled = False
            lblStatus.Text = "🔄 Rulez analiza..."

            modGlobals.Global_DBPath = txtPath.Text
            modGlobals.Global_UseCache = chkCache.IsChecked
            modGlobals.Global_UseParallel = chkParallel.IsChecked
            modGlobals.Global_VerboseLevel = If(rbQuiet.IsChecked, 0, If(rbVerbose.IsChecked, 2, 1))

            ' === rulează analiza în fundal ===
            Await Task.Run(Sub() VBA_TOKENIZER.Program.RunFullAnalysis())

            lblStatus.Text = "✅ Analiză completă!"

            ' 🔹 Lansează Viewer-ul, dar ascunde fereastra curentă (nu o închide)
            Dim viewer As New VBA_VIEWER.VBA_VIEWER.MainWindow()
            viewer.Tag = Me   ' salvăm referința la launcher
            viewer.Show()
            Me.Hide()
        End Sub

        Private Sub HandleLogMessage(msg As String, level As Integer)
            If modGlobals.Global_VerboseLevel = 0 Then Exit Sub

            Dispatcher.Invoke(Sub()
                                  AppendLog(msg)
                                  ' ⚡️ Forțează refresh vizual imediat
                                  txtLog.ScrollToEnd()
                                  txtLog.UpdateLayout()
                              End Sub, System.Windows.Threading.DispatcherPriority.Background)
        End Sub

        Private Sub AppendLog(msg As String)
            txtLog.AppendText(msg & vbCrLf)
        End Sub
    End Class
End Namespace
