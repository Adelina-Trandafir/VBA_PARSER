'PROJECT NAME: VBA_TOKENIZER
'FILE DESCRIPTION: Module pentru parsarea fișierelor de cod și header exportate din Access
'PATH: VBA_TOKENIZER/modFileParser.vb

Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports VBA_CORE
Imports VBA_CORE.Logger
Imports VBA_CORE.customTypes

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

            LogInfoLocal($"   → Total: {GlobalModules.Count + GlobalFormsReports.Count} surse încărcate.", 1)
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
            ElseIf detectedType = "Form" Then
                GlobalFormsReports.TryAdd(key, New FormReportContainer With {
                    .Name = modName,
                    .Type = detectedType,
                    .CodeContent = codeText,
                    .FilePath = cf
                })
            ElseIf detectedType = "Report" Then
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
        Dim validFormHeaders = headerFiles.Where(Function(hf)
                                                     Dim dirType = New DirectoryInfo(Path.GetDirectoryName(hf)).Name
                                                     Return dirType.Equals("Forms", StringComparison.OrdinalIgnoreCase)
                                                 End Function).ToList()

        Dim validReportHeaders = headerFiles.Where(Function(hf)
                                                       Dim dirType = New DirectoryInfo(Path.GetDirectoryName(hf)).Name
                                                       Return dirType.Equals("Reports", StringComparison.OrdinalIgnoreCase)
                                                   End Function).ToList()
        ' Opțiuni de paralelizare
        Dim opts As New ParallelOptions With {.MaxDegreeOfParallelism = Environment.ProcessorCount}

        ' Procesare paralelă sau secvențială
        If modGlobals.Global_UseParallel Then
            Parallel.ForEach(validFormHeaders, opts, Sub(hf) ProcessFormReportHeader(hf, "Form"))
            Parallel.ForEach(validReportHeaders, opts, Sub(hf) ProcessFormReportHeader(hf, "Report"))
        Else
            For Each hf In validFormHeaders
                ProcessFormReportHeader(hf, "Form")
            Next
            For Each hf In validReportHeaders
                ProcessFormReportHeader(hf, "Report")
            Next
        End If

        LogInfoLocal($" parsed sucessfully.", 1)
    End Sub

    Public Function PreProcessFormReportHeader(lines() As String) As FormReportDesigner
        Dim root As New FormReportDesigner()
        Dim stack As New Stack(Of FormReportDesigner)()
        stack.Push(root)

        Dim workingLines As New List(Of String)(lines)
        Dim i As Integer = 0
        Dim firstToken As String = ""
        Dim secondToken As String = ""

        ' Obiect temporar pentru proprietăți multi-linie
        Dim tempKey As String = Nothing
        Dim tempValues As List(Of String) = Nothing

        While i < workingLines.Count
            Dim line As String = workingLines(i)
            If line.Contains("CodeBehindForm") Then Exit While

            Dim trimmed As String = line.Trim()

            If Not String.IsNullOrWhiteSpace(trimmed) Then
                Dim matches As MatchCollection = RegexCache.RxDesigner.Matches(trimmed)

                If matches.Count > 0 Then
                    firstToken = matches(0).Value
                    secondToken = ""

                    If matches.Count = 3 Then
                        secondToken = matches(2).Value
                    ElseIf matches.Count = 2 Then
                        secondToken = matches(1).Value
                    End If

                    'Debug.Print($"{firstToken}:{secondToken}")

                    ' BEGIN e primul token → creează copil nou
                    If firstToken.Equals("Begin", StringComparison.OrdinalIgnoreCase) Then
                        Dim newNode As New FormReportDesigner() With {.RawName = secondToken}
                        If String.IsNullOrEmpty(newNode.Name) Then newNode.RawName = $"Unnamed_{stack.Peek().Children.Count}"
                        If secondToken.Equals("report", StringComparison.CurrentCultureIgnoreCase) Or secondToken.Equals("form", StringComparison.CurrentCultureIgnoreCase) Then
                            newNode.IsControl = True
                            newNode.IsSection = False
                        End If
                        stack.Peek().Children.Add(newNode)
                        stack.Push(newNode)

                    ElseIf firstToken.Equals("End", StringComparison.OrdinalIgnoreCase) Then
                        ' finalizează proprietate multi-linie
                        If tempKey IsNot Nothing AndAlso tempValues IsNot Nothing Then
                            stack.Peek().Properties.Add(tempKey, tempValues.ToArray())
                            tempKey = Nothing
                            tempValues = Nothing
                        ElseIf stack.Count > 1 Then
                            stack.Pop()
                        End If

                    Else
                        ' Dacă suntem în proprietate multi-linie, adaugă linia
                        If tempKey IsNot Nothing AndAlso tempValues IsNot Nothing Then
                            tempValues.Add(trimmed)
                        ElseIf firstToken.Trim().StartsWith(Chr(34)) AndAlso firstToken.EndsWith(Chr(34)) Then
                            If stack.Peek().Properties.Count > 0 Then
                                stack.Peek().Properties.AddLineToProperty(trimmed)
                            End If
                        Else
                            ' Proprietate nouă
                            Dim equalIndex As Integer = trimmed.IndexOf("="c)

                            If equalIndex > 0 Then
                                Dim key As String = trimmed.Substring(0, equalIndex).Trim()
                                Dim value As String = trimmed.Substring(equalIndex + 1).Trim()

                                ' Proprietate multi-linie?
                                If value.Equals("Begin", StringComparison.OrdinalIgnoreCase) Then
                                    tempKey = key
                                    tempValues = New List(Of String)()
                                Else
                                    stack.Peek().Properties.Add(key, {value})
                                End If
                            Else
                                stack.Peek().Properties.Add(trimmed, {})
                            End If
                        End If
                    End If
                End If
            End If
            i += 1
        End While
        Debug.Print(root.GetAllControls.Count)
        Return root
    End Function

    ''' <summary>
    ''' Parsează un fișier "_header.txt" de raport și construiește atât:
    ''' 1️⃣ arborele complet FormReportDesigner (root)
    ''' 2️⃣ containerul compatibil ReportContainer (Sections + Controls)
    ''' </summary>
    Private Sub ProcessFormReportHeader(hf As String, objType As String)
        Try
            Dim dirType = New DirectoryInfo(Path.GetDirectoryName(hf)).Name
            Dim formOrReportName = Path.GetFileNameWithoutExtension(hf).Replace("_header", "")
            Dim key = $"{formOrReportName}:{objType}"
            Dim lines = File.ReadAllLines(hf, Encoding.UTF8)

            ' === 1️⃣ Creează arborele logic FormReportDesigner ===
            Dim root As FormReportDesigner = PreProcessFormReportHeader(lines)

            ' === 2️⃣ Stochează arborele brut ===
            GlobalDesigners.TryAdd(key, root)

            ' === 3️⃣ Creează containerul tradițional ===
            Dim fci As FormReportContainer = Nothing
            If Not GlobalFormsReports.TryGetValue(key, fci) Then
                fci = New FormReportContainer With {
                    .FilePath = hf,
                    .Name = formOrReportName,
                    .Type = objType,
                    .DesignerContainer = root.GetControl(objType)
                }
                GlobalFormsReports.TryAdd(key, fci)
            End If

            ' === 4️⃣ Propietăți generale de raport ===
            fci.DesignerContainer = root.GetControl(objType, True)
            fci.RecordSource = fci.DesignerContainer.GetProperty("RecordSource", True)
            fci.Tag = fci.DesignerContainer.GetProperty("Tag", True, vbCrLf)

            Dim designerEvents = fci.DesignerContainer.GetAllPropertiesByValue("[Event Procedure]")

            If designerEvents.Count > 0 Then
                For Each prp In designerEvents
                    Dim evt As New EventSet With {
                        .EventName = prp.Key,
                        .HandlerString = fci.DesignerContainer.Name,
                        .RefObject = fci,
                        .RefType = fci.Type
                    }

                    fci.EventSets.Add(Strings.Join({fci.DesignerContainer.Name, IIf(prp.Key.StartsWith("On"), prp.Key.Substring(2), prp.Key)}, "_"), evt)
                Next
            End If

            ' === 5️⃣ Construiește toate secțiunile și controalele din arbore ===
            For Each secNode In fci.DesignerContainer.GetAllSections 'GetAllControls().Where(Function(f) f.IsSection)
                ' 5.1 Creează sau recuperează secțiunea asociată
                Dim sec As FormReportSection = Nothing
                If secNode IsNot Nothing Then
                    Dim secName As String = secNode.GetProperty("Name")
                    If String.IsNullOrEmpty(secName) Then secName = secNode.RawName

                    sec = fci.Sections.FirstOrDefault(Function(s) s.Name = secName)
                    If sec Is Nothing Then
                        sec = New FormReportSection With {
                            .Name = secName,
                            .Type = secNode.RawName,
                            .Parent = fci,
                            .DesignerControl = secNode
                        }
                        fci.Sections.Add(sec)
                    End If
                End If

                Dim sectionEvents = sec.DesignerControl.GetAllPropertiesByValue("[Event Procedure]")
                If sectionEvents.Count > 0 Then
                    For Each prp In sectionEvents
                        Dim evt As New EventSet With {
                            .EventName = prp.Key,
                            .HandlerString = sec.Name,
                            .RefType = "Section",
                            .RefObject = sec
                        }
                        fci.EventSets.Add(Strings.Join({sec.Name, IIf(prp.Key.StartsWith("On"), prp.Key.Substring(2), prp.Key)}, "_"), evt)
                    Next
                End If

                For Each ctrlNode In secNode.GetAllControls
                    ' 5.2 Creează controlul propriu-zis
                    Dim ctrl As New FormReportControl With {
                        .Type = ctrlNode.RawName,
                        .Name = ctrlNode.GetProperty("Name"),
                        .ControlSource = ctrlNode.GetProperty("ControlSource"),
                        .RecordSource = ctrlNode.GetProperty("RecordSource", True),
                        .RowSource = ctrlNode.GetProperty("RowSource"),
                        .SourceObject = ctrlNode.GetProperty("SourceObject"),
                        .Caption = ctrlNode.GetProperty("Caption"),
                        .Section = sec,
                        .DesignerControl = ctrlNode
                    }

                    sec.Controls.Add(ctrl)
                    ctrl.ParentSection = sec.Name
                    fci.Controls.Add(ctrl)

                    Dim controlEvents = ctrl.DesignerControl.GetAllPropertiesByValue("[Event Procedure]")

                    If controlEvents.Count > 0 Then
                        For Each prp In controlEvents
                            Dim evt As New EventSet With {
                                .EventName = prp.Key,
                                .HandlerString = ctrl.Name,
                                .RefObject = ctrl,
                                .RefType = "Control"
                            }

                            fci.EventSets.Add(Strings.Join({ctrl.Name, IIf(prp.Key.StartsWith("On"), prp.Key.Substring(2), prp.Key)}, "_"), evt)
                        Next
                    End If

                Next
            Next

            LogInfoLocal($"Parsed report header: {formOrReportName}", 1)

        Catch ex As Exception
            LogError($"ProcessReportHeader({Path.GetFileName(hf)})", ex)
        End Try
    End Sub

    Private Sub LogInfoLocal(ByVal msg As String, level As Integer, Optional force As Boolean = False)
        If pCurrentControl_FileParser.Value <> "" Then
            msg &= $"({pCurrentType_FileParser.Value}) [{pCurrentModule_FileParser.Value}]![{pCurrentControl_FileParser.Value}] : {msg}"
        Else
            msg &= $"({pCurrentType_FileParser.Value}) [{pCurrentModule_FileParser.Value}] : {msg}"
        End If

        LogInfo("[FILEPARSER]" & msg, level, force)
    End Sub
End Module