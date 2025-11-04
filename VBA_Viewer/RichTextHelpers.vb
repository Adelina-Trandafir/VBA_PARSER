Imports System.Text
Imports VBA_CORE
Imports VBA_CORE.modTypes
Imports VBA_VIEWER.VBA_VIEWER

Module RichTextHelpers
    Friend Sub WriteLinesToRTFDocument(mainWindow As MainWindow, sb As StringBuilder)
        If sb Is Nothing OrElse sb.Length = 0 Then
            mainWindow.txtSubDetail.Document = New FlowDocument(New Paragraph(New Run("(no details)")))
            Return
        End If

        Dim lines = sb.ToString().Split({vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
        If lines.Length = 0 Then
            mainWindow.txtSubDetail.Document = New FlowDocument(New Paragraph(New Run("(no details)")))
            Return
        End If

        Dim doc As New FlowDocument() With {
        .PagePadding = New Thickness(8),
        .FontFamily = New FontFamily("Consolas"),
        .FontSize = 13
    }

        Dim table As Table = Nothing
        Dim group As TableRowGroup = Nothing

        ' helper: începe un tabel nou
        Dim startTable = Sub()
                             table = New Table() With {.CellSpacing = 0, .Margin = New Thickness(0)}
                             ' ✅ coloana 1 fixă (ajustează după nevoie), coloana 2 ia restul
                             table.Columns.Add(New TableColumn() With {.Width = New GridLength(180)})
                             table.Columns.Add(New TableColumn() With {.Width = GridLength.Auto})
                             group = New TableRowGroup()
                             table.RowGroups.Add(group)
                         End Sub

        For Each raw In lines
            Dim line = raw.TrimEnd()
            If line.Length = 0 Then Continue For

            Dim prefix As String = ""
            Dim content As String = line

            If line.StartsWith("~~") Then
                prefix = "~~" : content = line.Substring(2).TrimStart()
            ElseIf line.StartsWith("~") Then
                prefix = "~" : content = line.Substring(1).TrimStart()
            ElseIf line.StartsWith("-") Then
                prefix = "-" : content = line.Substring(1).TrimStart()
            ElseIf line.StartsWith("=") Then
                prefix = "=" : content = line.Substring(1).TrimStart()
            End If

            Select Case prefix
                Case "~" ' TITLU
                    If table IsNot Nothing Then doc.Blocks.Add(table)
                    table = Nothing : group = Nothing

                    Dim p As New Paragraph(New Run(content) With {
                    .FontWeight = FontWeights.Bold,
                    .FontSize = 14,
                    .Foreground = Brushes.ForestGreen
                }) With {.Margin = New Thickness(0, 6, 0, 6), .LineHeight = Double.NaN}
                    doc.Blocks.Add(p)

                    startTable()

                Case Else
                    If table Is Nothing Then startTable()

                    Dim keyText As String = ""
                    Dim valText As String = ""

                    If content.Contains(":") Then
                        Dim parts = content.Split({":"c}, 2, StringSplitOptions.None)
                        keyText = parts(0).Trim()
                        If parts.Length > 1 Then valText = parts(1).Trim()
                    Else
                        keyText = content
                    End If

                    Dim row As New TableRow()

                    ' Celula KEY (bold, fără wrap, fixă)
                    Dim keyPara As New Paragraph(New Run(keyText) With {
                    .FontWeight = FontWeights.Bold,
                    .Foreground = Brushes.DarkCyan
                }) With {.Margin = New Thickness(0), .Padding = New Thickness(2, 0, 8, 0), .LineHeight = Double.NaN}
                    row.Cells.Add(New TableCell(keyPara) With {.Padding = New Thickness(0)})

                    ' Celula VALUE (wrap cu overflow, fără spații suplimentare)
                    Dim valBlock As New TextBlock() With {
                    .Text = valText,
                    .TextWrapping = TextWrapping.WrapWithOverflow,
                    .TextTrimming = TextTrimming.None
                }
                    Dim valBui As New BlockUIContainer(valBlock) With {.Margin = New Thickness(0)}
                    Dim valCell As New TableCell(valBui) With {.Padding = New Thickness(2, 0, 0, 0)}
                    row.Cells.Add(valCell)

                    group.Rows.Add(row)
            End Select
        Next

        doc.Blocks.Add(table)
        mainWindow.txtSubDetail.Document = doc
    End Sub

    Friend Function BuildTokenGridDocument(mainWindow As MainWindow, ml As MethodLine) As FlowDocument
        Dim doc As New FlowDocument() With {.PagePadding = New Thickness(0)}
        doc.Blocks.Clear()
        If ml?.Tokens Is Nothing OrElse ml.Tokens.Count = 0 Then
            doc.Blocks.Add(New Paragraph(New Run("(no tokens)")))
            Return doc
        End If

        ' ⭐ CURĂȚĂ dicționarul înainte de a popula grid-ul
        mainWindow.TokenReferences.Clear()

        ' === Grid principal ===
        Dim grid As New Grid() With {
        .ShowGridLines = False,
        .Margin = New Thickness(0),
        .HorizontalAlignment = HorizontalAlignment.Stretch
    }

        ' 4 coloane egale
        For i = 1 To 4
            grid.ColumnDefinitions.Add(New ColumnDefinition() With {.Width = New GridLength(1, GridUnitType.Star)})
        Next

        ' === Header ===
        Dim headerLabels = {"Name", "Type", "Method Pos", "Global Pos"}
        For i = 0 To headerLabels.Length - 1
            Dim border As New Border() With {
            .BorderBrush = Brushes.Gray,
            .BorderThickness = New Thickness(0.5),
            .Background = Brushes.LightGray,
            .Padding = New Thickness(4)
        }
            Dim tb As New TextBlock(New Run(headerLabels(i))) With {
            .FontWeight = FontWeights.Bold,
            .TextAlignment = TextAlignment.Center,
            .HorizontalAlignment = HorizontalAlignment.Center
        }
            border.Child = tb
            Grid.SetRow(border, 0)
            Grid.SetColumn(border, i)
            grid.Children.Add(border)
        Next
        grid.RowDefinitions.Add(New RowDefinition() With {.Height = GridLength.Auto})

        Dim currentRow As Integer = 1

        ' === Rânduri ===
        For Each t In ml.Tokens
            Dim dict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            For Each pair In t.ToString().Split(";"c)
                If pair.Contains("="c) Then
                    Dim kv = pair.Split({"="c}, 2, StringSplitOptions.None)
                    dict(kv(0).Trim()) = kv(1).Trim()
                End If
            Next

            grid.RowDefinitions.Add(New RowDefinition() With {.Height = GridLength.Auto})

            ' ⭐ CELULA PENTRU NAME (col 0) - CLICKABLE
            If dict.ContainsKey("Token") Then
                Dim border As New Border() With {
                .BorderBrush = Brushes.Gray,
                .BorderThickness = New Thickness(0.5),
                .Padding = New Thickness(4)
            }
                Dim value = dict("Token")

                ' ⭐ Creează un ID unic și salvează token-ul în dicționar
                Dim tokenId As String = Guid.NewGuid().ToString()
                mainWindow.TokenReferences(tokenId) = t

                Dim link As New Hyperlink(New Run(value)) With {
                .Tag = tokenId  ' ⭐ PUNE DOAR ID-ul STRING
            }
                AddHandler link.Click, AddressOf mainWindow.TokenName_Click
                Dim tb As New TextBlock()
                tb.Inlines.Add(link)
                border.Child = tb
                Grid.SetRow(border, currentRow)
                Grid.SetColumn(border, 0)
                grid.Children.Add(border)
            End If

            ' helper pentru celelalte celule
            Dim addCell = Sub(col As Integer, key As String)
                              If dict.ContainsKey(key) Then
                                  Dim border As New Border() With {
                                  .BorderBrush = Brushes.Gray,
                                  .BorderThickness = New Thickness(0.5),
                                  .Padding = New Thickness(4)
                              }
                                  Dim value = dict(key)
                                  Dim tb As New TextBlock(New Run(value))
                                  border.Child = tb
                                  Grid.SetRow(border, currentRow)
                                  Grid.SetColumn(border, col)
                                  grid.Children.Add(border)
                              End If
                          End Sub

            addCell(1, "Type")
            addCell(2, "Method_Pos")
            addCell(3, "Global_Pos")

            currentRow += 1
        Next

        doc.PagePadding = New Thickness(0)
        doc.Blocks.Add(New BlockUIContainer(grid) With {.Margin = New Thickness(0)})
        Return doc
    End Function

    Friend Sub ShowRTBTokenPropertiesPopup(token As Token)
        Dim popup As New Window()
        popup.Title = $"Token Properties: {token.TokenString}"
        popup.Width = 700
        popup.Height = 600
        popup.WindowStartupLocation = WindowStartupLocation.CenterScreen

        Dim rtb As New RichTextBox() With {
            .FontFamily = New FontFamily("Consolas"),
            .FontSize = 12,
            .IsReadOnly = True,
            .VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
            .Padding = New Thickness(10),
            .Background = New SolidColorBrush(Color.FromRgb(30, 30, 30))
        }

        ' Generează JSON
        Dim json = token.DumpJson(maxDepth:=2)
        Dim prettyJson = FormatJsonPretty(json)

        ' Aplică syntax highlighting
        ApplyJsonSyntaxHighlighting(rtb, prettyJson)

        Dim btnCopy As New Button() With {
            .Content = "📋 Copy JSON",
            .Width = 100,
            .Height = 30,
            .Margin = New Thickness(5)
        }
        AddHandler btnCopy.Click, Sub()
                                      Clipboard.SetText(prettyJson)
                                      MessageBox.Show("JSON copiat în clipboard!", "Success", MessageBoxButton.OK, MessageBoxImage.Information)
                                  End Sub

        Dim btnClose As New Button() With {
            .Content = "Close",
            .Width = 80,
            .Height = 30,
            .Margin = New Thickness(5)
        }
        AddHandler btnClose.Click, Sub() popup.Close()

        ' Layout buttons
        Dim buttonPanel As New StackPanel() With {
            .Orientation = Orientation.Horizontal,
            .HorizontalAlignment = HorizontalAlignment.Center,
            .Margin = New Thickness(10)
        }
        buttonPanel.Children.Add(btnCopy)
        buttonPanel.Children.Add(btnClose)

        Dim grid As New Grid()
        grid.RowDefinitions.Add(New RowDefinition() With {.Height = New GridLength(1, GridUnitType.Star)})
        grid.RowDefinitions.Add(New RowDefinition() With {.Height = GridLength.Auto})

        Grid.SetRow(rtb, 0)
        Grid.SetRow(buttonPanel, 1)
        grid.Children.Add(rtb)
        grid.Children.Add(buttonPanel)

        popup.Content = grid
        popup.ShowDialog()
    End Sub

    Private Sub ApplyJsonSyntaxHighlighting(rtb As RichTextBox, json As String)
        Dim doc As New FlowDocument()
        doc.PagePadding = New Thickness(0)
        doc.LineStackingStrategy = LineStackingStrategy.BlockLineHeight

        Dim para As New Paragraph()
        para.Margin = New Thickness(0)
        para.Padding = New Thickness(0)
        para.LineHeight = 8
        para.LineStackingStrategy = LineStackingStrategy.BlockLineHeight

        ' Culori pentru theme dark
        Dim colorBrace As Color = Color.FromRgb(220, 220, 220)      ' {, }, [, ] - alb
        Dim colorKey As Color = Color.FromRgb(156, 220, 254)        ' Property names - albastru deschis
        Dim colorString As Color = Color.FromRgb(206, 145, 120)     ' String values - portocaliu-maro
        Dim colorNumber As Color = Color.FromRgb(181, 206, 168)     ' Numbers - verde deschis
        Dim colorBoolean As Color = Color.FromRgb(86, 156, 214)     ' true/false/null - albastru
        Dim colorSymbol As Color = Color.FromRgb(180, 180, 180)     ' : , - gri
        Dim colorEmoji As Color = Color.FromRgb(255, 215, 0)        ' Emojis - galben

        ' Caractere care opresc un token
        Dim breakChars As String = "{}[],: " & vbCr & vbLf

        Dim i As Integer = 0
        Dim inString As Boolean = False
        Dim inKey As Boolean = False
        Dim afterColon As Boolean = False
        Dim escape As Boolean = False

        While i < json.Length
            Dim c As Char = json(i)

            ' Handle escape
            If escape Then
                para.Inlines.Add(New Run(c.ToString()) With {.Foreground = New SolidColorBrush(colorString)})
                escape = False
                i += 1
                Continue While
            End If

            If c = "\"c AndAlso inString Then
                para.Inlines.Add(New Run(c.ToString()) With {.Foreground = New SolidColorBrush(colorString)})
                escape = True
                i += 1
                Continue While
            End If

            Select Case c
                Case """"c
                    ' Quote - NU afișăm ghilimelele
                    If Not inString Then
                        ' Starting a string
                        inString = True
                        inKey = Not afterColon
                    Else
                        ' Ending a string
                        inString = False
                        inKey = False
                    End If

                Case ":"c
                    para.Inlines.Add(New Run(": ") With {.Foreground = New SolidColorBrush(colorSymbol)})
                    afterColon = True

                Case ","c
                    para.Inlines.Add(New Run(",") With {.Foreground = New SolidColorBrush(colorSymbol)})
                    afterColon = False

                Case "{"c, "}"c, "["c, "]"c
                    para.Inlines.Add(New Run(c.ToString()) With {.Foreground = New SolidColorBrush(colorBrace), .FontWeight = FontWeights.Bold})
                    afterColon = False

                Case vbLf, vbCr
                    para.Inlines.Add(New LineBreak())
                    afterColon = False

                Case " "c, vbTab
                    para.Inlines.Add(New Run(c.ToString()))

                Case Else
                    If inString Then
                        ' Inside string
                        Dim stringColor = If(inKey, colorKey, colorString)

                        ' Check for emojis (simple detection)
                        If AscW(c) > 127 Then
                            para.Inlines.Add(New Run(c.ToString()) With {.Foreground = New SolidColorBrush(colorEmoji)})
                        Else
                            para.Inlines.Add(New Run(c.ToString()) With {.Foreground = New SolidColorBrush(stringColor)})
                        End If
                    Else
                        ' Outside string - numbers, booleans, null
                        Dim token As String = ""
                        Dim startPos As Integer = i

                        While i < json.Length AndAlso Not breakChars.Contains(json(i))
                            token &= json(i)
                            i += 1
                        End While
                        i -= 1 ' Back one step

                        ' Determine token type
                        Dim tokenColor As Color
                        If token = "true" OrElse token = "false" OrElse token = "null" Then
                            tokenColor = colorBoolean
                        ElseIf IsNumeric(token) Then
                            tokenColor = colorNumber
                        Else
                            tokenColor = colorString
                        End If

                        para.Inlines.Add(New Run(token) With {.Foreground = New SolidColorBrush(tokenColor)})
                    End If
            End Select

            i += 1
        End While

        doc.Blocks.Add(para)
        rtb.Document = doc
    End Sub

    Private Function FormatJsonPretty(json As String) As String
        Dim sb As New StringBuilder()
        Dim indent As Integer = 0
        Dim inString As Boolean = False
        Dim escape As Boolean = False

        For i As Integer = 0 To json.Length - 1
            Dim c As Char = json(i)

            If escape Then
                sb.Append(c)
                escape = False
                Continue For
            End If

            If c = "\"c Then
                sb.Append(c)
                escape = True
                Continue For
            End If

            If c = """"c Then
                sb.Append(c)
                inString = Not inString
                Continue For
            End If

            If inString Then
                sb.Append(c)
                Continue For
            End If

            Select Case c
                Case "{"c, "["c
                    sb.Append(c)
                    sb.AppendLine()
                    indent += 2
                    sb.Append(New String(" "c, indent))

                Case "}"c, "]"c
                    sb.AppendLine()
                    indent -= 2
                    sb.Append(New String(" "c, indent))
                    sb.Append(c)

                Case ","c
                    sb.Append(c)
                    sb.AppendLine()
                    sb.Append(New String(" "c, indent))

                Case ":"c
                    sb.Append(c)
                    sb.Append(" ")

                Case " "c, vbTab, vbCr, vbLf
                    ' Skip

                Case Else
                    sb.Append(c)
            End Select
        Next

        Return sb.ToString()
    End Function
End Module
