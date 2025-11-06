'PROJECT NAME: VBA_TOKENIZER
'FILE DESCRIPTION: Module pentru exportul codului VBA dintr-o bază de date Access
'PATH: VBA_TOKENIZER/modExporter.vb
Imports Microsoft.Office.Interop.Access
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports VBA_CORE.Logger
Imports VBA_CORE
Imports System.Windows.Forms
Imports System.Web.UI

Public Module modExporter
    Private errorFiles As New List(Of String)

    ' ==========================================================
    '  EXPORT DIN ACCESS
    ' ==========================================================
    Public Sub ExportAccess() '(exported As List(Of ExportedModule))
        Dim sw As Stopwatch = Stopwatch.StartNew()
        LogInfo("=== Export Access VBA ===")
        Directory.CreateDirectory(modGlobals.Global_ExportDir)

        Dim app As Microsoft.Office.Interop.Access.Application = Nothing
        Try
            app = New Microsoft.Office.Interop.Access.Application()
            app.OpenCurrentDatabase(modGlobals.Global_DBPath)
            LogInfo("Access database opened...")

            Dim folders As New Dictionary(Of String, AcObjectType) From {
                {"Modules", AcObjectType.acModule},
                {"Forms", AcObjectType.acForm},
                {"Reports", AcObjectType.acReport},
                {"Macros", AcObjectType.acMacro}
            }

            For Each kvp In folders
                Dim folder = Path.Combine(modGlobals.Global_ExportDir, kvp.Key)
                Directory.CreateDirectory(folder)
                Dim objs As Object = Nothing
                Try
                    Select Case kvp.Value
                        Case AcObjectType.acModule : objs = app.CurrentProject.AllModules
                        Case AcObjectType.acForm : objs = app.CurrentProject.AllForms
                        Case AcObjectType.acReport : objs = app.CurrentProject.AllReports
                        Case AcObjectType.acMacro : objs = app.CurrentProject.AllMacros
                    End Select
                Catch ex As Exception
                    LogError($"List {kvp.Key}", ex)
                    Continue For
                End Try

                Dim count As Integer = 0
                For Each item In objs
                    Try
                        Dim name As String = item.Name
                        Dim safeName As String = SafeFileComponent(name)
                        Dim outFile = Path.Combine(folder, $"{safeName}.txt")
                        Dim fType As String = kvp.Key.TrimEnd("s"c)

                        app.SaveAsText(kvp.Value, name, outFile)
                        ExtractBinarySections(outFile, fType)

                        count += 1
                        LogInfo($"   [{count}] {kvp.Key}\{name}", 1)
                    Catch ex As Exception
                        LogError($"Export {kvp.Key}", ex)
                    End Try
                    Thread.Sleep(10)
                    Windows.Forms.Application.DoEvents()
                Next
            Next
        Catch ex As Exception
            LogError("ExportAccess", ex)
        Finally
            Try : app.CloseCurrentDatabase() : Catch : End Try
            Try : app.Quit() : Catch : End Try
            If app IsNot Nothing Then
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(app) : Catch : End Try
            End If
            app = Nothing

            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try

        sw.Stop()
        LogInfo($"Export finalized in {sw.Elapsed.TotalSeconds:N1}s")
    End Sub

    ' ==========================================================
    '  PRELOAD FILES
    ' ==========================================================
    'Public Sub PreloadAllFiles()
    '    fileCache.Clear()
    '    mapCache.Clear()

    '    Dim txtFiles = Directory.GetFiles(EXPORT_DIR, "*_code.txt", SearchOption.AllDirectories)
    '    Dim maps = Directory.GetFiles(EXPORT_DIR, "*_maps.txt", SearchOption.AllDirectories)

    '    LogInfo($"Preload: {txtFiles.Length} fișiere code + {maps.Length} maps")
    '    Dim opts As New ParallelOptions With {.MaxDegreeOfParallelism = Environment.ProcessorCount}

    '    Dim actionPerFile As Action(Of String) =
    '    Sub(p)
    '        Try
    '            Dim content = File.ReadAllText(p, Encoding.UTF8)
    '            Dim rawLines = content.Split({vbCrLf, vbLf}, StringSplitOptions.None).ToList()
    '            Dim merged As New List(Of String)
    '            Dim buffer As String = ""
    '            Dim lineNum As Integer = 1, startNum As Integer = 1

    '            For Each raw In rawLines
    '                Dim L = raw.TrimEnd()
    '                If L = "" OrElse L.StartsWith("'") Then
    '                    lineNum += 1
    '                    Continue For
    '                End If

    '                If buffer = "" Then startNum = lineNum
    '                If L.EndsWith("_") Then
    '                    buffer &= " " & L.TrimEnd("_"c, "&"c).Trim()
    '                Else
    '                    buffer &= " " & L.Trim()
    '                    merged.Add($"{startNum}: {buffer.Trim()}")
    '                    buffer = ""
    '                End If
    '                lineNum += 1
    '            Next
    '            If buffer <> "" Then merged.Add($"{startNum}: {buffer.Trim()}")

    '            Dim clean = merged.Select(Function(x) Regex.Replace(x, "^\d+:\s*", "")).ToArray()
    '            Dim pf As New PreloadedFile With {.Path = p, .Content = content, .Lines = merged.ToArray(), .CleanLines = clean}

    '            SyncLock fileCache
    '                fileCache(p) = pf
    '            End SyncLock
    '        Catch ex As Exception
    '            LogError($"Preload({Path.GetFileName(p)})", ex)
    '        End Try
    '    End Sub

    '    If useParallel Then
    '        Parallel.ForEach(txtFiles, opts, actionPerFile)
    '    Else
    '        For Each f In txtFiles : actionPerFile(f) : Next
    '    End If

    '    For Each m In maps
    '        Try
    '            mapCache(m) = File.ReadAllText(m, Encoding.UTF8)
    '        Catch ex As Exception
    '            LogError($"PreloadMap({Path.GetFileName(m)})", ex)
    '        End Try
    '    Next
    'End Sub

    ' ==========================================================
    '  EXTRACT HEADER / CODE
    ' ==========================================================
    Private Sub ExtractBinarySections(filePath As String, moduleType As String)
        Try
            Dim allText As String = File.ReadAllText(filePath, Encoding.UTF8)
            Dim lines = allText.Split({vbCrLf, vbLf}, StringSplitOptions.None)
            Dim optIndex As Integer = -1
            Dim isFormOrReport As Boolean = (moduleType = "Form" Or moduleType = "Report")

            If isFormOrReport Then
                optIndex = Array.FindIndex(lines, Function(L) L.ToLower().TrimStart().StartsWith("codebehindform", StringComparison.OrdinalIgnoreCase))
                If optIndex > -1 Then optIndex += 1
            End If

            Dim baseDir = Path.GetDirectoryName(filePath)
            Dim baseName = Path.GetFileNameWithoutExtension(filePath)
            Dim headerFile = Path.Combine(baseDir, baseName & "_header.txt")
            Dim codeFile = Path.Combine(baseDir, baseName & "_code.txt")

            If optIndex = -1 AndAlso isFormOrReport Then
                File.WriteAllText(headerFile, allText, Encoding.UTF8)
                LogInfo($"(i) {Path.GetFileName(filePath)} no code module.")
            ElseIf isFormOrReport Then
                Dim headerText = String.Join(vbCrLf, lines.Take(optIndex))
                Dim codeText = String.Join(vbCrLf, {$"'@@{moduleType}"}.Concat(lines.Skip(optIndex)))
                File.WriteAllText(headerFile, headerText, Encoding.UTF8)
                File.WriteAllText(codeFile, codeText, Encoding.UTF8)
            Else
                Dim codeText = String.Join(vbCrLf, {$"'@@{moduleType}"}.Concat(lines))
                File.WriteAllText(codeFile, codeText, Encoding.UTF8)
            End If

        Catch ex As Exception
            LogError($"ExtractBinarySections({Path.GetFileName(filePath)})", ex)
        End Try
    End Sub

    'Public Sub AttachCodeToFormReports()
    '    Program.LogInfo("=== AttachCodeToFormReports ===")

    '    Dim codeFiles = Directory.GetFiles(Program.EXPORT_DIR, "*_code.txt", SearchOption.AllDirectories)
    '    Dim count As Integer = 0

    '    For Each cf In codeFiles
    '        Try
    '            Dim dirType = New DirectoryInfo(Path.GetDirectoryName(cf)).Name
    '            If Not (dirType.Equals("Forms", StringComparison.OrdinalIgnoreCase) OrElse
    '                dirType.Equals("Reports", StringComparison.OrdinalIgnoreCase)) Then Continue For

    '            Dim formName = Path.GetFileNameWithoutExtension(cf).Replace("_code", "")
    '            Dim codeText = File.ReadAllText(cf, Encoding.UTF8)

    '            Dim fci As FormReportContainer = Nothing
    '            If GlobalFormsReports.TryGetValue(formName, fci) Then
    '                fci.CodeContent = codeText
    '            Else
    '                ' Formul n-a fost găsit în header — creează unul nou minimal
    '                fci = New FormReportContainer With {
    '                    .Name = formName,
    '                    .CodeContent = codeText
    '                    }
    '                GlobalFormsReports(formName) = fci
    '            End If

    '            count += 1
    '        Catch ex As Exception
    '            LogError($"AttachCodeToFormReports({Path.GetFileName(cf)})", ex)
    '        End Try
    '    Next

    '    Program.LogInfo($"   → {count} fișiere code atașate la Form/Report.")
    'End Sub

    Private Function SafeFileComponent(name As String) As String
        Dim s = Regex.Replace(name, "[^\w\- ]+", "_").Trim()
        If String.IsNullOrWhiteSpace(s) Then s = "unnamed"
        Return s
    End Function
End Module
