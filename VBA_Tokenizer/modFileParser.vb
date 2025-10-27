Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports VBA_CORE
Imports VBA_CORE.Logger
Imports VBA_CORE.modTypes
Imports VBA_CORE.VBA_CORE

Module modFileParser
    Public Sub CreateObjectsFromFiles()
        Try
            GlobalModules.Clear()

            ' Căutăm toate *_code.txt din Modules, Classes etc.
            Dim baseDirs = {"Modules", "Forms", "Reports"}
            Dim totalCount As Integer = 0

            For Each dirName In baseDirs
                Dim dirPath = Path.Combine(modGlobals.Global_ExportDir, dirName)
                If Not Directory.Exists(dirPath) Then Continue For

                For Each cf In Directory.GetFiles(dirPath, "*_code.txt", SearchOption.TopDirectoryOnly)
                    Try
                        Dim lines = File.ReadAllLines(cf, Encoding.UTF8)
                        Dim firstLine As String = If(lines.Length > 0, lines(0).Trim(), "")
                        Dim detectedType As String = "Module"

                        ' Detectăm tipul din header (ex: '@@Module', '@@Class')
                        If firstLine.StartsWith("'@@", StringComparison.OrdinalIgnoreCase) Then
                            detectedType = firstLine.TrimStart("'"c, "@"c).Trim()
                        End If

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

                        Dim modName = Path.GetFileNameWithoutExtension(cf).Replace("_code", "")
                        Dim codeText = File.ReadAllText(cf, Encoding.UTF8)
                        'LogInfo($"   → Loaded {detectedType}: {modName}")

                        If detectedType = "Module" Or detectedType = "Class" Then
                            GlobalModules.Add(String.Join(":", modName, detectedType), New ModuleContainer With {
                                .Name = modName,
                                .Type = detectedType,
                                .CodeContent = codeText,
                                .FilePath = cf,
                                .Lines = BuildMethodLinesFromCode(codeText)
                            })
                        Else
                            GlobalFormsReports.Add(String.Join(":", modName, detectedType), New FormReportContainer With {
                                .Name = modName,
                                .Type = detectedType,
                                .CodeContent = codeText,
                                .FilePath = cf,
                                .Lines = BuildMethodLinesFromCode(codeText)
                            })
                        End If

                        totalCount += 1
                    Catch ex As Exception
                        LogError($"BuildGlobalModulesFromCodeFiles({Path.GetFileName(cf)})", ex)
                        Stop
                    End Try
                Next
            Next

            'LogInfo($"   → GlobalModules: {totalCount} surse încărcate.")
        Catch ex As Exception
            LogError("BuildGlobalModulesFromCodeFiles", ex)
        End Try
    End Sub

    ' ==========================================================
    '  PARSE FORM / REPORT HEADERS
    ' ==========================================================
    Public Sub CreateControlsFromFiles()
        LogInfo("=== ParseFormReportHeaders (structură Access completă) ===")
        'GlobalFormsReports.Clear()

        Dim headerFiles = Directory.GetFiles(modGlobals.Global_ExportDir, "*_header.txt", SearchOption.AllDirectories)
        Dim inControlBlock As Boolean = False
        Dim controlDepth As Integer = 0

        For Each hf In headerFiles
            Try
                Dim dirType = New DirectoryInfo(Path.GetDirectoryName(hf)).Name

                If Not (dirType.Equals("Forms", StringComparison.OrdinalIgnoreCase) OrElse dirType.Equals("Reports", StringComparison.OrdinalIgnoreCase)) Then Continue For

                Dim objectType As String = IIf(dirType = "Forms", "Form", "Report")

                Dim formName = Path.GetFileNameWithoutExtension(hf).Replace("_header", "")

                Dim key = String.Join(":", formName, objectType)
                Dim fci As FormReportContainer = Nothing

                If GlobalFormsReports.ContainsKey(key) Then
                    fci = GlobalFormsReports(key)
                Else
                    'LogInfo($"Form/Report '{formName}' nu a fost găsit în GlobalFormsReports.")
                    fci = New FormReportContainer With {
                        .FilePath = hf,
                        .Name = formName,
                        .Type = objectType
                    }
                    GlobalFormsReports(key) = fci  ' adaugă automat lipsa
                End If

                'LogInfo($"Prelucrez header pentru {String.Join(":", formName, objectType)}")

                Dim lines = File.ReadAllLines(hf, Encoding.UTF8)

                ' frame de stivă cu referință la secțiunea curentă
                Dim stack As New Stack(Of (Type As String, Name As String, Body As List(Of String), Sec As FormSection))()

                Dim ignorePrtBlock As Boolean = False

                ' seturi de tipuri
                Dim rxSectionType As New Regex("^(Section|FormHeader|FormFooter|ReportHeader|ReportFooter|PageHeader|PageFooter|FormPageHeader|FormPageFooter|ReportPageHeader|ReportPageFooter)$", RegexOptions.IgnoreCase)
                Dim rxFormOrReport As New Regex("^(Form|Report)$", RegexOptions.IgnoreCase)


                For Each raw In lines
                    Dim L = raw.Trim()
                    If L = "" Then Continue For

                    ' ===== ignoră complet blocurile Prt... (ex: PrtDevMode = Begin ... End) =====
                    If ignorePrtBlock Then
                        If L.Equals("End", StringComparison.OrdinalIgnoreCase) Then ignorePrtBlock = False
                        Continue For
                    End If
                    If Regex.IsMatch(L, "(?i)=\s*Begin") Then
                        ignorePrtBlock = True
                        Continue For
                    End If

                    ' ===== dacă suntem în interiorul unui control =====
                    If inControlBlock Then
                        ' adaugăm linia în corpul controlului curent
                        stack.Peek().Body.Add(L)

                        ' detectăm dacă avem alte Begin/End imbricate (label/guid etc.)
                        If L.StartsWith("Begin ", StringComparison.OrdinalIgnoreCase) Then
                            controlDepth += 1
                        ElseIf L.Equals("End", StringComparison.OrdinalIgnoreCase) Then
                            controlDepth -= 1
                            If controlDepth <= 0 Then
                                ' ieșim din modul control
                                inControlBlock = False

                                ' finalizează controlul
                                Dim block = stack.Pop()
                                Dim blockText As String = String.Join(vbCrLf, block.Body)

                                ' dacă e control valid și are secțiune
                                If block.Sec IsNot Nothing Then
                                    Dim mName = RegexCache.RxName.Match(blockText)
                                    If mName.Success Then
                                        Dim ctrl As New FormControl With {
                                            .Type = block.Type,
                                            .Name = CleanPropValue(mName.Groups(1).Value),
                                            .Section = block.Sec
                                        }

                                        ParseSingleControl(ctrl, block.Body, fci)

                                        fci.Controls.Add(ctrl)
                                        block.Sec.Controls.Add(ctrl)
                                    End If
                                End If

                                ' adaugăm conținutul în blocul părinte
                                If stack.Count > 0 Then
                                    stack.Peek().Body.AddRange(block.Body)
                                End If
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

                        ElseIf rxSectionType.IsMatch(t) Then
                            Dim secName As String = n
                            Dim fs As New FormSection With {.Type = t, .Name = secName, .Parent = fci}
                            fci.Sections.Add(fs)
                            stack.Push((t, secName, New List(Of String), fs))

                        ElseIf t <> "" Then
                            ' === am intrat într-un control ===
                            stack.Push((t, "", New List(Of String), parentSec))
                            inControlBlock = True
                            controlDepth = 1
                        Else
                            ' container generic
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

                'Stop
                'SyncLock GlobalFormsReports
                '    GlobalFormsReports(formName) = fci
                'End SyncLock
            Catch ex As Exception
                LogError($"ParseFormReportHeaders({Path.GetFileName(hf)})", ex)
            End Try
        Next

        LogInfo($"   → {GlobalFormsReports.Count} form/report header(e) analizate.")
    End Sub

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
                        ' nu includ linia "Begin ..." în labelLines (nu e nevoie)
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

                ' Construiește sub-labelul și asociază-l o singură dată
                If fc.SubLabel Is Nothing Then
                    Dim lbl As New FormControl With {.Type = "Label", .Section = fc.Section}
                    ParseSingleControl(lbl, labelLines, fci)

                    fc.SubLabel = lbl

                    ' adaugă labelul și în secțiune, dacă nu există deja
                    If fc.Section IsNot Nothing AndAlso Not fc.Section.Controls.Any(Function(c) c.Name.Equals(lbl.Name, StringComparison.OrdinalIgnoreCase)) Then
                        fc.Section.Controls.Add(lbl)
                    End If

                    ' adaugă labelul și în lista globală de controale a formularului
                    If fci IsNot Nothing AndAlso Not fci.Controls.Any(Function(c) c.Name.Equals(lbl.Name, StringComparison.OrdinalIgnoreCase)) Then
                        fci.Controls.Add(lbl)
                    End If
                End If

                ' continuă cu linia curentă (deja avansată după End)
                Continue While
            End If

            ' === Parsare proprietăți ale controlului curent (doar dacă regex-ul se potrivește) ===
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
    '  PARSE FUNCȚII + EXISTING EXPORTS
    ' ==========================================================
    'Public Sub ParseFunctionDefinitions(ByRef m As ExportedModule)
    '    Try
    '        Dim content = File.ReadAllText(m.FilePath)
    '        Dim lines = content.Split({vbCrLf, vbLf}, StringSplitOptions.None)
    '        Dim rx As New Regex("^(?<acc>Public|Private)?\s*(?<kind>Sub|Function|Property|Class|End\s+Class)\s*(?<name>[A-Za-z0-9_]+)?", RegexOptions.IgnoreCase Or RegexOptions.Multiline)
    '        Dim matches = rx.Matches(content)

    '        Dim defs As New List(Of (idx As Integer, sig As String, kind As String, name As String, acc As String))
    '        For Each mm As Match In matches
    '            Dim sig = mm.Value.Trim()
    '            Dim nm = If(mm.Groups("name").Success, mm.Groups("name").Value.Trim(), "")
    '            Dim kind = mm.Groups("kind").Value.Trim()
    '            Dim acc = If(mm.Groups("acc").Success, mm.Groups("acc").Value, "Public")
    '            Dim startIdx As Integer = content.Substring(0, mm.Index).Split({vbCrLf, vbLf}, StringSplitOptions.None).Length
    '            defs.Add((startIdx, sig, kind, nm, acc))
    '        Next

    '        Dim currentClass As String = ""
    '        For i = 0 To defs.Count - 1
    '            Dim def = defs(i)
    '            If def.kind.Equals("Class", StringComparison.OrdinalIgnoreCase) Then
    '                currentClass = def.name : Continue For
    '            End If
    '            If def.kind.Equals("End Class", StringComparison.OrdinalIgnoreCase) Then
    '                currentClass = "" : Continue For
    '            End If
    '            If {"Sub", "Function", "Property"}.Contains(def.kind, StringComparer.OrdinalIgnoreCase) Then
    '                Dim fi As New FunctionInfo With {
    '                    .Signature = def.sig,
    '                    .Accessibility = def.acc,
    '                    .ParentClass = currentClass,
    '                    .Name = If(currentClass <> "", $"{currentClass}.{def.name}", def.name),
    '                    .StartLine = def.idx
    '                }
    '                Dim endLine As Integer = If(i < defs.Count - 1, defs(i + 1).idx - 1, lines.Length)
    '                fi.EndLine = Math.Max(fi.StartLine, endLine)
    '                fi.LineCount = Math.Max(0, fi.EndLine - fi.StartLine + 1)
    '                Dim cplx As Integer = lines.Skip(fi.StartLine - 1).Take(fi.LineCount).Count(Function(L) Regex.IsMatch(L, "\b(If|ElseIf|For|Do|While|Select|Case|Until)\b", RegexOptions.IgnoreCase))
    '                fi.Complexity = cplx
    '                m.Functions.Add(fi)
    '            End If
    '        Next
    '    Catch ex As Exception
    '        LogError($"ParseFunctionDefinitions({m.Name})", ex)
    '    End Try
    'End Sub


    ' === helper: curăță valorile proprietăților (elimină ghilimele, trim) ===
    Private Function CleanPropValue(v As String) As String
        Dim s = v.Trim()
        If s.StartsWith("""") AndAlso s.EndsWith("""") AndAlso s.Length >= 2 Then
            s = s.Substring(1, s.Length - 2)
        End If
        Return s.Trim()
    End Function
End Module
