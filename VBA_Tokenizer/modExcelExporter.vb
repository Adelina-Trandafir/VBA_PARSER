Imports System.Collections.Concurrent
Imports System.ComponentModel
Imports System.IO
Imports System.Text
Imports VBA_CORE.modTypes
Imports VBA_CORE
Imports VBA_CORE.Logger
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Module modExcelExporter

    ' Thread-safe queue pentru a colecta datele în paralel
    Private ReadOnly ExportQueue As New ConcurrentQueue(Of ExportRow)

    ' Structură pentru o linie de export
    Public Class ExportRow
        Public Property Level0 As String  ' Module Name
        Public Property Level1 As String  ' Type (Module/Form/Report)
        Public Property Level2 As String  ' Member Type (Method/Property/Variable/etc)
        Public Property Level3 As String  ' Member Name
        Public Property Level4 As String  ' Property Name
        Public Property Level5 As String  ' Property Value
        Public Property ModuleName As String  ' Pentru sortare
        Public Property SortOrder As Integer  ' Pentru a păstra ordinea logică
    End Class

    ''' <summary>
    ''' Adaugă date din ParseSingleCodeBlock la coada de export
    ''' THREAD-SAFE: poate fi apelat din Parallel.ForEach
    ''' </summary>
    Public Sub CollectModuleData(moduleName As String, moduleType As String, formOrModule As Object)
        Try
            Dim sortBase As Integer = ExportQueue.Count * 1000

            ' === Proprietăți modulului ===
            ExportQueue.Enqueue(New ExportRow With {
                .Level0 = moduleName,
                .Level1 = moduleType,
                .Level2 = "Property",
                .Level3 = "Name",
                .Level4 = "",
                .Level5 = moduleName,
                .ModuleName = moduleName,
                .SortOrder = sortBase + 1
            })

            ExportQueue.Enqueue(New ExportRow With {
                .Level0 = moduleName,
                .Level1 = moduleType,
                .Level2 = "Property",
                .Level3 = "Type",
                .Level4 = "",
                .Level5 = moduleType,
                .ModuleName = moduleName,
                .SortOrder = sortBase + 2
            })

            ' === Type-specific exports ===
            If TypeOf formOrModule Is ModuleContainer Then
                ExportModuleContainer(CType(formOrModule, ModuleContainer), sortBase + 100)
            ElseIf TypeOf formOrModule Is FormReportContainer Then
                ExportFormReportContainer(CType(formOrModule, FormReportContainer), sortBase + 100)
            End If

            'LogInfo($"CollectModuleData({moduleName}) - date colectate.")
        Catch ex As Exception
            LogError($"CollectModuleData({moduleName})", ex)
        End Try
    End Sub

    Private Sub ExportModuleContainer(m As ModuleContainer, sortBase As Integer)
        Dim idx As Integer = sortBase

        ' === Constants ===
        If m.Constants IsNot Nothing AndAlso m.Constants.Count > 0 Then
            For Each c In m.Constants
                ExportQueue.Enqueue(New ExportRow With {
                    .Level0 = m.Name,
                    .Level1 = m.Type,
                    .Level2 = "Constant",
                    .Level3 = c.Name,
                    .Level4 = "Type",
                    .Level5 = c.Type,
                    .ModuleName = m.Name,
                    .SortOrder = idx
                })
                idx += 1

                If Not String.IsNullOrEmpty(c.Value) Then
                    ExportQueue.Enqueue(New ExportRow With {
                        .Level0 = m.Name,
                        .Level1 = m.Type,
                        .Level2 = "Constant",
                        .Level3 = c.Name,
                        .Level4 = "Value",
                        .Level5 = c.Value,
                        .ModuleName = m.Name,
                        .SortOrder = idx
                    })
                    idx += 1
                End If
            Next
        End If

        ' === Variables ===
        If m.Variables IsNot Nothing AndAlso m.Variables.Count > 0 Then
            For Each v In m.Variables
                ExportQueue.Enqueue(New ExportRow With {
                    .Level0 = m.Name,
                    .Level1 = m.Type,
                    .Level2 = "Variable",
                    .Level3 = v.Name,
                    .Level4 = "Type",
                    .Level5 = v.Type,
                    .ModuleName = m.Name,
                    .SortOrder = idx
                })
                idx += 1

                ExportQueue.Enqueue(New ExportRow With {
                    .Level0 = m.Name,
                    .Level1 = m.Type,
                    .Level2 = "Variable",
                    .Level3 = v.Name,
                    .Level4 = "IsGlobal",
                    .Level5 = v.IsGlobal.ToString(),
                    .ModuleName = m.Name,
                    .SortOrder = idx
                })
                idx += 1

                ExportQueue.Enqueue(New ExportRow With {
                    .Level0 = m.Name,
                    .Level1 = m.Type,
                    .Level2 = "Variable",
                    .Level3 = v.Name,
                    .Level4 = "IsPublic",
                    .Level5 = v.IsPublic.ToString(),
                    .ModuleName = m.Name,
                    .SortOrder = idx
                })
                idx += 1
            Next
        End If

        ' === Declares ===
        If m.Declares IsNot Nothing AndAlso m.Declares.Count > 0 Then
            For Each kvp In m.Declares
                Dim decl = kvp.Value
                ExportMethodInfo(m.Name, m.Type, "Declare", decl, idx)
                idx += 100
            Next
        End If

        ' === Methods ===
        If m.Methods IsNot Nothing AndAlso m.Methods.Count > 0 Then
            For Each kvp In m.Methods
                Dim method = kvp.Value
                ExportMethodInfo(m.Name, m.Type, "Method", method, idx)
                idx += 100
            Next
        End If
    End Sub

    Private Sub ExportFormReportContainer(f As FormReportContainer, sortBase As Integer)
        Dim idx As Integer = sortBase

        ' === RecordSource ===
        If Not String.IsNullOrEmpty(f.RecordSource) Then
            ExportQueue.Enqueue(New ExportRow With {
                .Level0 = f.Name,
                .Level1 = f.Type,
                .Level2 = "Property",
                .Level3 = "RecordSource",
                .Level4 = "",
                .Level5 = f.RecordSource,
                .ModuleName = f.Name,
                .SortOrder = idx
            })
            idx += 1
        End If

        ' === Tag ===
        If Not String.IsNullOrEmpty(f.Tag) Then
            ExportQueue.Enqueue(New ExportRow With {
                .Level0 = f.Name,
                .Level1 = f.Type,
                .Level2 = "Property",
                .Level3 = "Tag",
                .Level4 = "",
                .Level5 = f.Tag,
                .ModuleName = f.Name,
                .SortOrder = idx
            })
            idx += 1
        End If

        ' === Constants ===
        If f.Constants IsNot Nothing AndAlso f.Constants.Count > 0 Then
            For Each c In f.Constants
                ExportQueue.Enqueue(New ExportRow With {
                    .Level0 = f.Name,
                    .Level1 = f.Type,
                    .Level2 = "Constant",
                    .Level3 = c.Name,
                    .Level4 = "Type",
                    .Level5 = c.Type,
                    .ModuleName = f.Name,
                    .SortOrder = idx
                })
                idx += 1
            Next
        End If

        ' === Variables ===
        If f.Variables IsNot Nothing AndAlso f.Variables.Count > 0 Then
            For Each v In f.Variables
                ExportQueue.Enqueue(New ExportRow With {
                    .Level0 = f.Name,
                    .Level1 = f.Type,
                    .Level2 = "Variable",
                    .Level3 = v.Name,
                    .Level4 = "Type",
                    .Level5 = v.Type,
                    .ModuleName = f.Name,
                    .SortOrder = idx
                })
                idx += 1
            Next
        End If

        ' === Sections ===
        If f.Sections IsNot Nothing AndAlso f.Sections.Count > 0 Then
            For Each sec In f.Sections
                ExportQueue.Enqueue(New ExportRow With {
                    .Level0 = f.Name,
                    .Level1 = f.Type,
                    .Level2 = "Section",
                    .Level3 = sec.Name,
                    .Level4 = "Type",
                    .Level5 = sec.Type,
                    .ModuleName = f.Name,
                    .SortOrder = idx
                })
                idx += 1

                ' Controls în secțiune
                If sec.Controls IsNot Nothing Then
                    For Each ctrl In sec.Controls
                        ExportControl(f.Name, f.Type, sec.Name, ctrl, idx)
                        idx += 50
                    Next
                End If
            Next
        End If

        ' === All Controls (flat) ===
        If f.Controls IsNot Nothing AndAlso f.Controls.Count > 0 Then
            For Each ctrl In f.Controls
                ExportControl(f.Name, f.Type, "AllControls", ctrl, idx)
                idx += 50
            Next
        End If

        ' === Declares ===
        If f.Declares IsNot Nothing AndAlso f.Declares.Count > 0 Then
            For Each kvp In f.Declares
                Dim decl = kvp.Value
                ExportMethodInfo(f.Name, f.Type, "Declare", decl, idx)
                idx += 100
            Next
        End If

        ' === Methods ===
        If f.Methods IsNot Nothing AndAlso f.Methods.Count > 0 Then
            For Each kvp In f.Methods
                Dim method = kvp.Value
                ExportMethodInfo(f.Name, f.Type, "Method", method, idx)
                idx += 100
            Next
        End If
    End Sub

    Private Sub ExportControl(moduleName As String, moduleType As String, sectionName As String, ctrl As FormControl, ByRef idx As Integer)
        ' Control header
        ExportQueue.Enqueue(New ExportRow With {
            .Level0 = moduleName,
            .Level1 = moduleType,
            .Level2 = $"Section[{sectionName}]",
            .Level3 = $"Control[{ctrl.Type}]",
            .Level4 = "Name",
            .Level5 = ctrl.Name,
            .ModuleName = moduleName,
            .SortOrder = idx
        })
        idx += 1

        If Not String.IsNullOrEmpty(ctrl.ControlSource) Then
            ExportQueue.Enqueue(New ExportRow With {
                .Level0 = moduleName,
                .Level1 = moduleType,
                .Level2 = $"Section[{sectionName}]",
                .Level3 = $"Control[{ctrl.Type}]",
                .Level4 = "ControlSource",
                .Level5 = ctrl.ControlSource,
                .ModuleName = moduleName,
                .SortOrder = idx
            })
            idx += 1
        End If

        If Not String.IsNullOrEmpty(ctrl.RecordSource) Then
            ExportQueue.Enqueue(New ExportRow With {
                .Level0 = moduleName,
                .Level1 = moduleType,
                .Level2 = $"Section[{sectionName}]",
                .Level3 = $"Control[{ctrl.Type}]",
                .Level4 = "RecordSource",
                .Level5 = ctrl.RecordSource,
                .ModuleName = moduleName,
                .SortOrder = idx
            })
            idx += 1
        End If

        If Not String.IsNullOrEmpty(ctrl.RowSource) Then
            ExportQueue.Enqueue(New ExportRow With {
                .Level0 = moduleName,
                .Level1 = moduleType,
                .Level2 = $"Section[{sectionName}]",
                .Level3 = $"Control[{ctrl.Type}]",
                .Level4 = "RowSource",
                .Level5 = ctrl.RowSource,
                .ModuleName = moduleName,
                .SortOrder = idx
            })
            idx += 1
        End If

        If Not String.IsNullOrEmpty(ctrl.Caption) Then
            ExportQueue.Enqueue(New ExportRow With {
                .Level0 = moduleName,
                .Level1 = moduleType,
                .Level2 = $"Section[{sectionName}]",
                .Level3 = $"Control[{ctrl.Type}]",
                .Level4 = "Caption",
                .Level5 = ctrl.Caption,
                .ModuleName = moduleName,
                .SortOrder = idx
            })
            idx += 1
        End If

        ' SubLabel recursiv
        If ctrl.SubLabel IsNot Nothing Then
            ExportControl(moduleName, moduleType, $"{sectionName}.{ctrl.Name}", ctrl.SubLabel, idx)
        End If
    End Sub

    Private Sub ExportMethodInfo(moduleName As String, moduleType As String, memberType As String, method As MethodInfo, ByRef idx As Integer)
        ' Method header
        ExportQueue.Enqueue(New ExportRow With {
            .Level0 = moduleName,
            .Level1 = moduleType,
            .Level2 = memberType,
            .Level3 = method.Name,
            .Level4 = "Signature",
            .Level5 = method.Signature,
            .ModuleName = moduleName,
            .SortOrder = idx
        })
        idx += 1

        ExportQueue.Enqueue(New ExportRow With {
            .Level0 = moduleName,
            .Level1 = moduleType,
            .Level2 = memberType,
            .Level3 = method.Name,
            .Level4 = "StartLine",
            .Level5 = method.StartLine.ToString(),
            .ModuleName = moduleName,
            .SortOrder = idx
        })
        idx += 1

        ExportQueue.Enqueue(New ExportRow With {
            .Level0 = moduleName,
            .Level1 = moduleType,
            .Level2 = memberType,
            .Level3 = method.Name,
            .Level4 = "EndLine",
            .Level5 = method.EndLine.ToString(),
            .ModuleName = moduleName,
            .SortOrder = idx
        })
        idx += 1

        ExportQueue.Enqueue(New ExportRow With {
            .Level0 = moduleName,
            .Level1 = moduleType,
            .Level2 = memberType,
            .Level3 = method.Name,
            .Level4 = "MethodScope",
            .Level5 = method.MethodScope,
            .ModuleName = moduleName,
            .SortOrder = idx
        })
        idx += 1

        If method.IsFunction Then
            ExportQueue.Enqueue(New ExportRow With {
                .Level0 = moduleName,
                .Level1 = moduleType,
                .Level2 = memberType,
                .Level3 = method.Name,
                .Level4 = "ReturnType",
                .Level5 = method.ReturnType,
                .ModuleName = moduleName,
                .SortOrder = idx
            })
            idx += 1
        End If

        ' Parameters
        If method.Parameters IsNot Nothing AndAlso method.Parameters.Count > 0 Then
            For Each p In method.Parameters
                ExportQueue.Enqueue(New ExportRow With {
                    .Level0 = moduleName,
                    .Level1 = moduleType,
                    .Level2 = memberType,
                    .Level3 = method.Name,
                    .Level4 = $"Param[{p.Name}]",
                    .Level5 = $"Type={p.Type}, ByRef={p.IsByRef}, Optional={p.IsOptional}",
                    .ModuleName = moduleName,
                    .SortOrder = idx
                })
                idx += 1
            Next
        End If

        ' Variables locale
        If method.Variables IsNot Nothing AndAlso method.Variables.Count > 0 Then
            For Each v In method.Variables
                ExportQueue.Enqueue(New ExportRow With {
                    .Level0 = moduleName,
                    .Level1 = moduleType,
                    .Level2 = memberType,
                    .Level3 = method.Name,
                    .Level4 = $"LocalVar[{v.Name}]",
                    .Level5 = $"Type={v.Type}",
                    .ModuleName = moduleName,
                    .SortOrder = idx
                })
                idx += 1
            Next
        End If
    End Sub

    ''' <summary>
    ''' Scrie toate datele acumulate în Excel
    ''' APELEAZĂ LA FINAL, după ce s-a terminat paralelizarea
    ''' </summary>
    Public Sub WriteToExcel(outputPath As String)
        Try
            LogInfo($"[Excel] Scriu {ExportQueue.Count} linii în {outputPath}...")

            ' Sortează datele
            Dim sortedData = ExportQueue.OrderBy(Function(r) r.ModuleName).ThenBy(Function(r) r.SortOrder).ToList()

            ' Setează licența EPPlus (necesară pentru versiuni recente)
            ExcelPackage.License.SetNonCommercialPersonal("Adelina Trandafir")

            Using package As New ExcelPackage()
                Dim ws = package.Workbook.Worksheets.Add("VBA Structure")

                ' Header
                ws.Cells(1, 1).Value = "Module Name"
                ws.Cells(1, 2).Value = "Type"
                ws.Cells(1, 3).Value = "Member Type"
                ws.Cells(1, 4).Value = "Member Name"
                ws.Cells(1, 5).Value = "Property Name"
                ws.Cells(1, 6).Value = "Property Value"

                ' Format header
                Using rng = ws.Cells(1, 1, 1, 6)
                    rng.Style.Font.Bold = True
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid
                    rng.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray)
                    rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thick
                End Using

                ' Scrie datele
                Dim row As Integer = 2
                For Each item In sortedData
                    ws.Cells(row, 1).Value = item.Level0
                    ws.Cells(row, 2).Value = item.Level1
                    ws.Cells(row, 3).Value = item.Level2
                    ws.Cells(row, 4).Value = item.Level3
                    ws.Cells(row, 5).Value = item.Level4
                    ws.Cells(row, 6).Value = item.Level5
                    row += 1
                Next

                ' Auto-fit columns
                ws.Cells.AutoFitColumns()

                ' Freeze header
                ws.View.FreezePanes(2, 1)

                ' Salvează
                package.SaveAs(New FileInfo(outputPath))
            End Using

            LogInfo($"✅ Excel salvat: {outputPath}")

        Catch ex As Exception
            LogError("WriteToExcel", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Resetează coada (opțional, dacă vrei să faci mai multe export-uri)
    ''' </summary>
    Public Sub ClearQueue()
        Dim dummy As ExportRow = Nothing
        While ExportQueue.TryDequeue(dummy)
        End While
    End Sub

End Module