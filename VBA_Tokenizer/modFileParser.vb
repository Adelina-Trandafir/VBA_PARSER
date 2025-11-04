'PROJECT NAME: VBA_TOKENIZER
'FILE DESCRIPTION: Module pentru parsarea fișierelor de cod și header exportate din Access
'PATH: VBA_TOKENIZER/modFileParser.vb

Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports VBA_CORE
Imports VBA_CORE.Logger
Imports VBA_CORE.modTypes

Module modFileParser
    ' ==========================================================
    '  BUILD OBJECTS FROM CODE FILES (cu suport paralelizare)
    ' ==========================================================
    Friend pCurrentModule_FileParser As New ThreadLocal(Of String)(Function() "")
    Friend pCurrentType_FileParser As New ThreadLocal(Of String)(Function() "")
    Friend pCurrentControl_FileParser As New ThreadLocal(Of String)(Function() "")

    Public Sub CreateObjectsFromFiles()
        Try
            GlobalModules.Clear()
            GlobalFormsReports.Clear()

            ' Colectăm toate fișierele *_code.txt
            Dim baseDirs = {"Modules", "Forms", "Reports"}
            Dim allFiles As New List(Of String)

            For Each dirName In baseDirs
                Dim dirPath = Path.Combine(modGlobals.Global_ExportDir, dirName)
                If Not Directory.Exists(dirPath) Then Continue For

                allFiles.AddRange(Directory.GetFiles(dirPath, "*_code.txt", SearchOption.TopDirectoryOnly))
            Next

            ' Opțiuni de paralelizare
            Dim opts As New ParallelOptions With {.MaxDegreeOfParallelism = Environment.ProcessorCount}

            ' Procesare paralelă sau secvențială
            If modGlobals.Global_UseParallel Then
                Parallel.ForEach(allFiles, opts, Sub(cf) ProcessSingleCodeFile(cf))
            Else
                For Each cf In allFiles
                    ProcessSingleCodeFile(cf)
                Next
            End If

            'LogInfo($"   → Total: {GlobalModules.Count + GlobalFormsReports.Count} surse încărcate.")
        Catch ex As Exception
            LogError("CreateObjectsFromFiles", ex)
        End Try
    End Sub

    ' ==========================================================
    '  PROCESARE UN SINGUR FIȘIER CODE (rulează paralel)
    ' ==========================================================
    Private Sub ProcessSingleCodeFile(cf As String)
        Try
            Dim lines = File.ReadAllLines(cf, Encoding.UTF8)
            Dim firstLine As String = If(lines.Length > 0, lines(0).Trim(), "")
            Dim detectedType As String = "Module"

            ' Detectăm tipul din header (ex: '@@Module', '@@Class')
            If firstLine.StartsWith("'@@", StringComparison.OrdinalIgnoreCase) Then
                detectedType = firstLine.TrimStart("'"c, "@"c).Trim()
            End If

            Try
                ' Normalizează tipul
                Select Case True
                    Case detectedType.Equals("class", StringComparison.OrdinalIgnoreCase)
                        detectedType = "Class"
                    Case detectedType.Equals("module", StringComparison.OrdinalIgnoreCase)
                        detectedType = "Module"
                    Case detectedType.Equals("form", StringComparison.OrdinalIgnoreCase)
                        detectedType = "Form"
                    Case detectedType.Equals("report", StringComparison.OrdinalIgnoreCase)
                        detectedType = "Report"
                    Case Else
                        Throw New Exception($"Tip nevalid detectat în fișier: {cf}:{detectedType}")
                End Select
            Catch ex As Exception
                LogError(detectedType, ex)
            End Try

            Dim modName = Path.GetFileNameWithoutExtension(cf).Replace("_code", "")
            Dim codeText = File.ReadAllText(cf, Encoding.UTF8)

            pCurrentModule_FileParser.Value = modName
            pCurrentType_FileParser.Value = detectedType

            LogInfoLocal("Started parsing.", 1)

            Dim key = String.Join(":", modName, detectedType)

            ' Adăugare thread-safe în ConcurrentDictionary
            If detectedType = "Module" Or detectedType = "Class" Then
                GlobalModules.TryAdd(key, New ModuleContainer With {
                    .Name = modName,
                    .Type = detectedType,
                    .CodeContent = codeText,
                    .FilePath = cf
                })
            Else
                GlobalFormsReports.TryAdd(key, New FormReportContainer With {
                    .Name = modName,
                    .Type = detectedType,
                    .CodeContent = codeText,
                    .FilePath = cf
                })
            End If

            LogInfoLocal($"Parsed successfully.", 1)

        Catch ex As Exception
            LogError($"ProcessSingleCodeFile({Path.GetFileName(cf)})", ex)
            'ConditionalStop()  ' Stop doar în modul secvențial (debug)
        End Try
    End Sub

    ' ==========================================================
    '  PARSE FORM / REPORT HEADERS (cu suport paralelizare)
    ' ==========================================================
    Public Sub CreateControlsFromFiles()
        LogInfoLocal("=== Parse Form/Report Headers ===", 0)

        Dim headerFiles = Directory.GetFiles(modGlobals.Global_ExportDir, "*_header.txt", SearchOption.AllDirectories).ToList()

        ' Filtrare pentru Forms și Reports
        Dim validHeaders = headerFiles.Where(Function(hf)
                                                 Dim dirType = New DirectoryInfo(Path.GetDirectoryName(hf)).Name
                                                 Return dirType.Equals("Forms", StringComparison.OrdinalIgnoreCase) OrElse
                                                        dirType.Equals("Reports", StringComparison.OrdinalIgnoreCase)
                                             End Function).ToList()

        ' Opțiuni de paralelizare
        Dim opts As New ParallelOptions With {.MaxDegreeOfParallelism = Environment.ProcessorCount}

        ' Procesare paralelă sau secvențială
        If modGlobals.Global_UseParallel Then
            Parallel.ForEach(validHeaders, opts, Sub(hf) ProcessSingleHeaderFile(hf))
        Else
            For Each hf In validHeaders
                ProcessSingleHeaderFile(hf)
            Next
        End If

        LogInfoLocal($" parsed sucessfully.", 1)
    End Sub

    ' ==========================================================
    '  PROCESARE UN SINGUR FIȘIER HEADER (rulează paralel)
    ' ==========================================================
    Private Sub ProcessSingleHeaderFile(hf As String)
        Try
            Dim dirType = New DirectoryInfo(Path.GetDirectoryName(hf)).Name
            Dim objectType As String = IIf(dirType = "Forms", "Form", "Report")
            Dim formName = Path.GetFileNameWithoutExtension(hf).Replace("_header", "")
            Dim key = String.Join(":", formName, objectType)

            ' Obține sau creează container
            Dim fci As FormReportContainer = Nothing
            If Not GlobalFormsReports.TryGetValue(key, fci) Then
                fci = New FormReportContainer With {
                    .FilePath = hf,
                    .Name = formName,
                    .Type = objectType
                }
                GlobalFormsReports.TryAdd(key, fci)
            End If

            ' Variabile locale de stare pentru parsing (thread-safe)
            Dim inControlBlock As Boolean = False
            Dim controlDepth As Integer = 0
            Dim stack As New Stack(Of (Type As String, Name As String, Body As List(Of String), Sec As FormSection))()
            Dim ignorePrtBlock As Boolean = False

            ' Cache regex-uri
            Dim rxSectionType = RegexCache.RxSectionType
            Dim rxFormOrReport = RegexCache.RxFormOrReport
            Dim rxPrtBegin = RegexCache.RxBeginPrtForm

            Dim lines = File.ReadAllLines(hf, Encoding.UTF8)

            For Each raw In lines
                Dim L = raw.Trim()
                If L = "" Then Continue For

                ' ===== ignoră complet blocurile Prt... =====
                If ignorePrtBlock Then
                    If L.Equals("End", StringComparison.OrdinalIgnoreCase) Then
                        ignorePrtBlock = False
                        LogInfoLocal("End of Prt... block", 2)
                    End If
                    Continue For
                End If
                If rxPrtBegin.Matches(L).Count > 0 Then
                    ignorePrtBlock = True
                    LogInfoLocal($"Prt Block {rxPrtBegin.Matches(L).Item(0).Groups(1).Value}: {L}", 2)
                    Continue For
                End If

                ' ===== dacă suntem în interiorul unui control =====
                If inControlBlock Then
                    stack.Peek().Body.Add(L)

                    If L.StartsWith("Begin ", StringComparison.OrdinalIgnoreCase) Then
                        controlDepth += 1
                        LogInfoLocal($"Nested Begin ({controlDepth}): {L}", 2)
                        pCurrentControl_FileParser.Value = stack.Peek().Name

                    ElseIf L.Equals("End", StringComparison.OrdinalIgnoreCase) Then
                        controlDepth -= 1
                        If controlDepth <= 0 Then
                            inControlBlock = False
                            pCurrentControl_FileParser.Value = ""

                            ' finalizează controlul
                            Dim controlBlock = stack.Pop()
                            Dim ctrlType As String = controlBlock.Type
                            Dim parentSec As FormSection = controlBlock.Sec
                            Dim ctrl = New FormControl With {.Type = ctrlType, .Section = parentSec}

                            ParseSingleControl(ctrl, controlBlock.Body, fci)

                            ' adaugă controlul în părinte
                            If parentSec IsNot Nothing Then
                                parentSec.Controls.Add(ctrl)
                                ctrl.ParentSection = parentSec.Name
                            End If

                            ' adaugă în lista globală dacă nu există deja
                            If Not fci.Controls.Any(Function(c) c.Name.Equals(ctrl.Name, StringComparison.OrdinalIgnoreCase)) Then
                                fci.Controls.Add(ctrl)
                            End If

                            LogInfoLocal($"Parsed Control Finished", 2)
                        End If
                    End If

                    Continue For
                End If

                ' ===== Begin ... =====
                If L.StartsWith("Begin ", StringComparison.OrdinalIgnoreCase) Then
                    Dim parts = L.Split({" "c}, StringSplitOptions.RemoveEmptyEntries)
                    Dim t As String = If(parts.Length > 1, parts(1), "")
                    Dim n As String = If(parts.Length > 2, parts(2), "")
                    Dim parentSec As FormSection = If(stack.Count > 0, stack.Peek().Sec, Nothing)

                    If rxFormOrReport.IsMatch(t) Then
                        stack.Push((t, formName, New List(Of String), Nothing))
                        LogInfoLocal($"Form/Report block.", 2)

                    ElseIf rxSectionType.IsMatch(t) Then
                        Dim secName As String = n
                        Dim fs As New FormSection With {.Type = t, .Name = secName, .Parent = fci}
                        fci.Sections.Add(fs)
                        stack.Push((t, secName, New List(Of String), fs))
                        LogInfoLocal($"Section block: {t} - {secName}", 2)

                    ElseIf t <> "" Then
                        stack.Push((t, "", New List(Of String), parentSec))
                        inControlBlock = True
                        controlDepth = 1
                        pCurrentControl_FileParser.Value = t

                        LogInfoLocal($"Control block started", 2)
                    Else
                        stack.Push(("(Group)", "", New List(Of String), parentSec))
                    End If

                    Continue For
                End If

                ' ===== End =====
                If L.Equals("End", StringComparison.OrdinalIgnoreCase) Then
                    If stack.Count = 0 Then Continue For
                    Dim block = stack.Pop()
                    Dim blockText As String = String.Join(vbCrLf, block.Body)

                    If rxFormOrReport.IsMatch(block.Type) Then
                        Dim mTag = RegexCache.RxTag.Match(blockText)
                        Dim mRs = RegexCache.RxRecSrc.Match(blockText)
                        If mTag.Success Then fci.Tag = CleanPropValue(mTag.Groups(1).Value)
                        If mRs.Success Then fci.RecordSource = CleanPropValue(mRs.Groups(1).Value)
                    End If

                    If stack.Count > 0 Then
                        stack.Peek().Body.AddRange(block.Body)
                    End If

                    Continue For
                End If

                ' ===== Linii obișnuite =====
                If stack.Count > 0 Then
                    stack.Peek().Body.Add(L)
                End If
            Next

            LogInfoLocal($"Parsed header successfully.", 1)

        Catch ex As Exception
            LogError($"ProcessSingleHeaderFile({Path.GetFileName(hf)})", ex)
            'ConditionalStop()  ' Stop doar în modul secvențial (debug)
        End Try
    End Sub

    ' ==========================================================
    '  HELPER: PARSE SINGLE CONTROL
    ' ==========================================================
    Private Sub ParseSingleControl(ByRef fc As FormControl, block As List(Of String), Optional fci As FormReportContainer = Nothing)
        Dim i As Integer = 0
        While i < block.Count
            Dim ln As String = block(i)

            ' === Detectează un bloc "Begin Label ... End" imbricat în control ===
            If Regex.IsMatch(ln, "^(?im)^\s*Begin\s+Label\b") Then
                Dim depth As Integer = 1
                Dim labelLines As New List(Of String)

                i += 1 ' intrăm după linia "Begin Label"
                While i < block.Count AndAlso depth > 0
                    Dim l2 As String = block(i)

                    If Regex.IsMatch(l2, "^(?im)^\s*Begin\b") Then
                        depth += 1
                        i += 1
                        Continue While
                    End If

                    If Regex.IsMatch(l2, "^(?im)^\s*End\s*$") Then
                        depth -= 1
                        i += 1
                        If depth = 0 Then Exit While
                        Continue While
                    End If

                    labelLines.Add(l2.Trim())
                    i += 1
                End While

                ' Construiește sub-labelul
                If fc.SubLabel Is Nothing Then
                    Dim lbl As New FormControl With {.Type = "Label", .Section = fc.Section}
                    ParseSingleControl(lbl, labelLines, fci)
                    fc.SubLabel = lbl

                    ' adaugă labelul în secțiune
                    If fc.Section IsNot Nothing AndAlso Not fc.Section.Controls.Any(Function(c) c.Name.Equals(lbl.Name, StringComparison.OrdinalIgnoreCase)) Then
                        fc.Section.Controls.Add(lbl)
                    End If

                    ' adaugă labelul în lista globală
                    If fci IsNot Nothing AndAlso Not fci.Controls.Any(Function(c) c.Name.Equals(lbl.Name, StringComparison.OrdinalIgnoreCase)) Then
                        fci.Controls.Add(lbl)
                    End If
                End If

                Continue While
            End If

            ' === Parsare proprietăți ale controlului curent ===
            Dim m As Match

            m = Regex.Match(ln, "(?im)^\s*Name\s*=\s*""?([^""]+)""?")
            If m.Success Then fc.Name = m.Groups(1).Value.Trim()

            m = Regex.Match(ln, "(?im)^\s*ControlSource\s*=\s*(.+)$")
            If m.Success Then fc.ControlSource = m.Groups(1).Value.Trim()

            m = Regex.Match(ln, "(?im)^\s*RecordSource\s*=\s*(.+)$")
            If m.Success Then fc.RecordSource = m.Groups(1).Value.Trim()

            m = Regex.Match(ln, "(?im)^\s*RowSource\s*=\s*(.+)$")
            If m.Success Then fc.RowSource = m.Groups(1).Value.Trim()

            m = Regex.Match(ln, "(?im)^\s*SourceObject\s*=\s*(.+)$")
            If m.Success Then fc.SourceObject = m.Groups(1).Value.Trim()

            m = Regex.Match(ln, "(?im)^\s*Caption\s*=\s*(.+)$")
            If m.Success Then fc.Caption = m.Groups(1).Value.Trim()

            ' evenimente [Event Procedure]
            Dim rxEvt As New Regex("(?im)^\s*(On\w+)\s*=\s*""\[Event Procedure\]""")
            For Each em As Match In rxEvt.Matches(ln)
                fc.Caption &= $" [Evt:{em.Groups(1).Value}]"
            Next

            i += 1
        End While
    End Sub

    ' ==========================================================
    '  HELPER: CLEAN PROPERTY VALUE
    ' ==========================================================
    Private Function CleanPropValue(v As String) As String
        Dim s = v.Trim()
        If s.StartsWith("""") AndAlso s.EndsWith("""") AndAlso s.Length >= 2 Then
            s = s.Substring(1, s.Length - 2)
        End If
        Return s.Trim()
    End Function

    Private Sub LogInfoLocal(ByVal msg As String, level As Integer, Optional force As Boolean = False)
        If pCurrentControl_FileParser.Value <> "" Then
            msg = $"({pCurrentType_FileParser.Value}) [{pCurrentModule_FileParser.Value}]![{pCurrentControl_FileParser.Value}] : {msg}"
        Else
            msg = $"({pCurrentType_FileParser.Value}) [{pCurrentModule_FileParser.Value}] : {msg}"
        End If

        LogInfo(msg, level, force)
    End Sub
End Module