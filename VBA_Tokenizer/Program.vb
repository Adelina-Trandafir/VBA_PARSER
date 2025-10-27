Imports AVACONT_Viewer
Imports AVACONT_Core
Imports AVACONT_Core.Logger
Imports System.IO

Public Module Program
    ' ===============================
    ' 🏁 MAIN
    ' ===============================
    <STAThread()>
    Sub Main()
        Directory.CreateDirectory(EXPORT_DIR)
        Dim swTotal As Stopwatch = Stopwatch.StartNew()
        Dim exported As New List(Of Object)

        ' ===== CLI ARGUMENTS =====
        Dim args = Environment.GetCommandLineArgs().Skip(1).Select(Function(a) a.Trim().ToLowerInvariant()).ToArray()
        If args.Contains("--nothreads") OrElse args.Contains("nothreads") OrElse args.Contains("debug") OrElse args.Contains("--debug") Then
            useParallel = False : LogInfo("⚙️  Rulez secvențial (fără Parallel.ForEach).", True)
            verboseLevel = 2
        End If
        If args.Contains("--quiet") OrElse args.Contains("quiet") Then
            verboseLevel = 0
        ElseIf args.Contains("--verbose") OrElse args.Contains("verbose") Then
            verboseLevel = 2
        End If
        Dim noOpen As Boolean = args.Contains("--noopen") OrElse args.Contains("noopen")
        Dim useCache As Boolean = args.Contains("--cache") OrElse args.Contains("cache")

        If useCache Then
            LogInfo("🔄 Mod CACHE activat - încerc să încarc din cache...")
            If modSerializer.LoadCache() Then
                LogInfo("✅ Date încărcate din cache!")

                ' Sari direct la UI
                Try
                    Dim app As New AVACONT_Viewer.App()
                    Dim win As New AVACONT_Viewer.AVACONT_Viewer.MainWindow()
                    app.Run(win)
                Catch ex As Exception
                    LogError("LaunchWPF", ex)
                End Try

                Exit Sub
            Else
                LogInfo("⚠️ Cache invalid sau inexistent. Procesez normal...")
            End If
        End If

        Try
            ' ===== EXPORT / ANALIZĂ EXISTENTĂ =====
            If noOpen Then
                'LogInfo("▶ Mod NO-OPEN: analizez fișierele deja exportate.", True)
                'modExporter.AnalyzeExistingExports(exported)
            Else
                ' 1) EXPORT din Access
                Dim swExport As Stopwatch = Stopwatch.StartNew()
                modExporter.ExportAccess()'(exported)
                swExport.Stop()
                LogInfo($"[Timer] Export Access: {swExport.Elapsed.TotalSeconds:N1}s")

                ' 2) ANALIZĂ fișiere exportate (include preload)
                'Dim swAna As Stopwatch = Stopwatch.StartNew()
                'modExporter.AnalyzeExistingExports(exported)
                'swAna.Stop()
                'LogInfo($"[Timer] AnalyzeExistingExports: {swAna.Elapsed.TotalSeconds:N1}s")
            End If

            ' ===== ANALIZĂ STRUCTURALĂ =====
            Dim swAn As Stopwatch = Stopwatch.StartNew()

            swAn.Restart()
            CreateObjectsFromFiles()
            LogInfo($"[Timer] CreateObjectsFromFiles: {swAn.Elapsed.TotalSeconds:N1}s")

            swAn.Restart()
            CreateControlsFromFiles()
            LogInfo($"[Timer] CreateControlsFromFiles: {swAn.Elapsed.TotalSeconds:N1}s")

            swAn.Restart()
            ParseTextToObjects()
            LogInfo($"[Timer] ParseTextToObjects: {swAn.Elapsed.TotalSeconds:N1}s")

            swAn.Restart()
            modSerializer.SaveCache()
            LogInfo($"[Timer] SaveCache: {swAn.Elapsed.TotalSeconds:N1}s")

            'swAn.Restart()
            'Dim excelPath = Path.Combine(EXPORT_DIR, "VBA_Structure.xlsx")
            'modExcelExporter.WriteToExcel(excelPath)
            'LogInfo($"[Timer] Excel export: {swAn.Elapsed.TotalSeconds:N1}s")

            ' ===== HEADER PARSING (Form / Report Controls) =====
            'modFileParser.CreateControlsFromFiles()
            'LogInfo($"[Timer] ParseFormReportHeaders: {swAn.Elapsed.TotalSeconds:N1}s")

            ' ===== ATTACH CODE TO FORM/REPORT MODULES =====
            'swAn.Restart()
            'modExporter.AttachCodeToFormReports()
            'LogInfo($"[Timer] AttachCodeToFormReports: {swAn.Elapsed.TotalSeconds:N1}s")

            ' ===== BUILD GlobalModules din Modules\*_code.txt =====
            swAn.Restart()
            'modExporter.BuildGlobalModulesFromCodeFiles()
            LogInfo($"[Timer] BuildGlobalModulesFromCodeFiles: {swAn.Elapsed.TotalSeconds:N1}s")

            ' ===== SYMBOLS & REFERENCES =====
            swAn.Restart()
            'modCodeParser.ParseTextToObjects()
            LogInfo($"[Timer] BuildSymbolDatabase: {swAn.Elapsed.TotalSeconds:N1}s")

            ' ===== CLASIFICARE TOKEN-URI & ANALIZA DEPENDENȚELOR =====
            swAn.Restart()
            'Analizer_Classifier.ClassifyTokensAndAnalyzeDependencies()
            LogInfo($"[Timer] ClassifyTokensAndAnalyzeDependencies: {swAn.Elapsed.TotalSeconds:N1}s")

            Try
                Dim app As New AVACONT_Viewer.App()
                Dim win As New AVACONT_Viewer.AVACONT_Viewer.MainWindow()
                app.Run(win)
            Catch ex As Exception
                LogError("LaunchWPF", ex)
            End Try

            ' ===== DEAD CODE =====
            'swAn.Restart()

            'Dim dead As New List(Of String)

            'For Each em In exported
            '    For Each f In em.Functions
            '        Dim isEvent As Boolean = modReferenceResolver.IsTypicalAccessEvent(f.Name) AndAlso
            '                             (em.Type = "Forms" OrElse em.Type = "Reports")
            '        If Not isEvent AndAlso f.CalledBy.Count = 0 Then
            '            dead.Add($"[{em.Type}] {em.Name} :: {f.Signature}")
            '        End If
            '    Next
            'Next
            'File.WriteAllLines(Path.Combine(EXPORT_DIR, "DeadCode.txt"), dead, Encoding.UTF8)
            'LogInfo($"DeadCode.txt: {dead.Count} funcții fără referințe.")

            '' ===== CIRCULAR DEPENDENCIES =====
            'Dim cycles = modReferenceResolver.DetectModuleCycles(exported)
            'File.WriteAllLines(Path.Combine(EXPORT_DIR, "CircularDependencies.txt"), cycles, Encoding.UTF8)
            'LogInfo($"CircularDependencies.txt: {cycles.Count} cicluri detectate.")

            'swAn.Stop()
            'LogInfo($"[Timer] Analiză simboluri+referințe: {swAn.Elapsed.TotalSeconds:N1}s")

            '' ===== RAPOARTE =====
            'Dim swRep As Stopwatch = Stopwatch.StartNew()
            'modReports.GenerateReports(exported)
            'modReports.GenerateHTML(exported)
            'modReports.ExportAIReady(exported)
            'swRep.Stop()
            'LogInfo($"[Timer] Rapoarte: {swRep.Elapsed.TotalSeconds:N1}s")

        Catch ex As Exception
            LogError("Main", ex)
            LogInfo("❌ Eroare generală: " & ex.Message, True)
        End Try

        swTotal.Stop()
        LogInfo($"✅ Gata. Total: {swTotal.Elapsed.TotalSeconds:N1}s", True)
    End Sub


End Module
