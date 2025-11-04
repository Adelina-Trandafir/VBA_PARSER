'PROJECT NAME: VBA_TOKENIZER
'FILE DESCRIPTION: Module principal pentru rularea analizei complete
'PATH: VBA_TOKENIZER/Program.vb

Imports System.IO
Imports VBA_CORE
Imports VBA_CORE.Logger

Public Module Program
    <STAThread()>
    Public Sub Main()
        ' doar pentru compatibilitate CLI
        RunFullAnalysis()
    End Sub

    <STAThread()>
    Public Sub RunFullAnalysis()
        Dim swTotal As Stopwatch = Stopwatch.StartNew()
        Directory.CreateDirectory(modGlobals.Global_ExportDir)

        Try
            LogInfo("=== Running VBA Analyzer ===", 0, True)

            ' ===== 1️⃣ Încarcă din cache dacă e selectat =====
            If modGlobals.Global_UseCache Then
                If modSerializer.LoadCache() Then
                    Exit Sub
                Else
                    LogInfo("⚠️ Error loading data from cache. Running full analysis...", 0, True)
                End If
            End If

            ' ===== 2️⃣ Dacă e selectat un fișier DB, îl procesez =====
            If File.Exists(modGlobals.Global_DBPath) Then
                LogInfo("📂 Openin ACCESS database: " & modGlobals.Global_DBPath, 0, True)
                modExporter.ExportAccess()
            Else
                LogInfo("⚠️ No ACCESS database provided! Trying local files (C:\ExportVBA)...", 0, True)
            End If

            ' ===== 3️⃣ Prelucrare logică =====
            CreateObjectsFromFiles()
            CreateControlsFromFiles()
            ParseTextToObjects()
            TokenizeAllModules()
            'CleanupWithReferences()
            ' Build relations
            'modSymbolResolutionRunner.RunFullSymbolResolution()

            ' ===== 4️⃣ Salvare cache =====
            'modSerializer.SaveCache()

            swTotal.Stop()
            LogInfo($"✅ Time taken to complete: {swTotal.Elapsed.TotalSeconds:N1}s", 0, True)

        Catch ex As Exception
            LogError("RunFullAnalysis", ex)
        End Try
    End Sub
End Module
