Imports System.Text.RegularExpressions
Imports System.Text
Imports VBA_CORE.modTypes
Imports VBA_CORE

Public Module Tokenizer

    ' ==============================================================
    ' 🔥  GENEREAZĂ METHODLINES dintr-un bloc de cod complet
    ' ==============================================================
    Public Function BuildMethodLinesFromCode(code As String, Optional context As String = "") As List(Of MethodLine)
        Try
            Dim result As New List(Of MethodLine)
            If String.IsNullOrWhiteSpace(code) Then Return result

            ' === 1️⃣ Normalizează codul (unește liniile și elimină & _ între stringuri)
            Dim normalizedLines = NormalizeAndNumberCode(code)

            ' === 2️⃣ Procesează fiecare linie normalizată
            For Each raw In normalizedLines
                ' Format: "index:startLine: text"
                Dim parts = raw.Split({":"c}, 3)
                If parts.Length < 3 Then Continue For

                Dim outNum As Integer = Val(parts(0))
                Dim srcStart As Integer = Val(parts(1))
                Dim text As String = parts(2).Trim()

                ' === Curăță pentru tokenizare (fără stringuri, comentarii, date)
                Dim clean = CleanLineForTokenization(text)

                ' === Extrage tokenii (TokenString)
                'Dim tokens = GetTokensFromLine(clean, srcStart, context)

                ' === Creează MethodLine
                Dim ml As New MethodLine With {
                .LineNumber = srcStart,
                .Content = text
            }

                result.Add(ml)
            Next
            Return result

        Catch ex As Exception
            LogError("BuildMethodLinesFromCode", ex)
            Return New List(Of MethodLine)
        End Try

    End Function

    ' ==============================================================
    ' 🔹 Normalizează și numerotează codul
    ' ==============================================================
    Public Function NormalizeAndNumberCode(code As String) As List(Of String)
        Dim result As New List(Of String)
        If String.IsNullOrWhiteSpace(code) Then Return result

        Dim rawLines = code.Split({vbCrLf, vbLf}, StringSplitOptions.None).ToList()

        Dim buffer As String = ""
        Dim lineNum As Integer = 1
        Dim startLine As Integer = 1

        For Each raw In rawLines
            Dim L = raw.TrimEnd()

            ' === 0️⃣ Sărim complet peste liniile speciale @@ (nu contează la numerotare)
            If L.TrimStart().StartsWith("'@@") Then
                Continue For
            End If

            ' === 1️⃣ Linii de tip Option, comentarii sau goale → contează la numărătoare, dar nu se includ
            If L.Trim() = "" OrElse
               L.TrimStart().StartsWith("'") OrElse
               L.TrimStart().StartsWith("Option", StringComparison.OrdinalIgnoreCase) Then
                lineNum += 1
                Continue For
            End If

            ' === 2️⃣ Dacă bufferul e gol, reținem prima linie de concatenare
            If buffer = "" Then startLine = lineNum

            ' === 3️⃣ Elimină concatenările stringurilor ("aaa" & _ "bbb" → "aaabbb")
            L = Regex.Replace(L, """\s*&\s*_+\s*""", "")
            L = Regex.Replace(L, "\s{2,}", " ").Trim()

            ' === 4️⃣ Continuă linia dacă se termină în "_"
            If L.EndsWith("_") Then
                buffer &= " " & L.TrimEnd("_"c, " "c, vbTab)
            Else
                buffer &= " " & L.Trim()
                result.Add($"{result.Count + 1}:{startLine}: {buffer.Trim()}")
                buffer = ""
            End If

            lineNum += 1
        Next

        ' === 5️⃣ Adaugă ultima linie (dacă e ceva rămas)
        If buffer <> "" Then
            result.Add($"{result.Count + 1}:{startLine}: {buffer.Trim()}")
        End If

        Return result
    End Function

    ' ==============================================================
    ' 🔹 Curăță linia de string-uri, comentarii, date etc.
    ' ==============================================================

    Public Function CleanLineForTokenization(line As String) As String
        If String.IsNullOrEmpty(line) Then Return ""
        Dim result As New StringBuilder(line.Length)
        Dim inString As Boolean = False
        Dim i As Integer = 0

        While i < line.Length
            Dim ch As Char = line(i)
            If ch = """"c Then
                If inString AndAlso i + 1 < line.Length AndAlso line(i + 1) = """"c Then
                    result.Append("  ")
                    i += 2
                    Continue While
                Else
                    inString = Not inString
                    result.Append(" "c)
                    i += 1
                    Continue While
                End If
            End If

            If inString Then
                result.Append(" "c)
            Else
                result.Append(ch)
            End If
            i += 1
        End While

        Dim clean = result.ToString()
        Dim commentPos = clean.IndexOf("'"c)
        If commentPos >= 0 Then clean = clean.Substring(0, commentPos)
        clean = RegexCache.RxDateLiteral.Replace(clean, " ")
        Return clean.Trim()
    End Function

    ' ==============================================================
    ' 🔹 Extrage tokeni dintr-o linie curățată
    ' ==============================================================
    Public Function GetTokensFromLine(cleanLine As String, lineNumber As Integer, methodContext As String) As List(Of Token)
        Dim tokens As New List(Of Token)
        Dim prevToken As Token = Nothing
        Dim prevOp As String = Nothing

        For Each m As Match In RegexCache.RxTokenizer.Matches(cleanLine)
            Dim name = m.Groups(1).Value
            Dim op = m.Groups(2).Value

            If name.Length = 0 Then Continue For
            If Char.IsDigit(name(0)) Then Continue For
            If RegexCache.VbaKeywords.Contains(name) Then Continue For

            Dim tType As String = "identifier"  ' DEFAULT
            Dim isLhs As Boolean = False
            Dim parentName As String = Nothing

            ' Parent din lanțul anterior
            If prevOp = "." AndAlso prevToken IsNot Nothing Then
                parentName = prevToken.TokenString
            End If

            ' LOGICĂ CONSISTENTĂ:
            Select Case op
                Case "."
                    tType = "object"     ' obiectul dinaintea punctului

                Case "("
                    If prevOp = "." Then
                        tType = "method"     ' metoda după punct: obj.Method(
                    Else
                        tType = "function"   ' funcție standalone: Function(
                    End If

                Case "="
                    If prevToken Is Nothing OrElse prevOp <> "." Then
                        tType = "variable"   ' variabilă LHS: var = 
                        isLhs = True
                    Else
                        tType = "property"   ' proprietate: obj.Prop = 
                    End If

                Case ")", "", vbCrLf, Nothing
                    If prevOp = "." Then
                        tType = "property"   ' proprietate fără paranteze: obj.Prop
                    ElseIf prevOp = "(" Then
                        tType = "parameter"  ' parametru în apel: Func(param)
                    Else
                        tType = "identifier" ' identificator simplu
                    End If

                Case ",", "+", "-", "*", "/", "&"
                    If prevOp = "(" Then
                        tType = "parameter"  ' parametru în listă: Func(p1, p2)
                    Else
                        tType = "operand"    ' operand în expresie: a + b
                    End If

                Case Else
                    tType = "identifier"
            End Select

            Dim tok As New Token With {
            .LineNumber = lineNumber,
            .TokenString = name,
            .NextSymbol = op,
            .TokenType = tType,
            .Parent = prevToken,
            .IsLeftSide = isLhs,
            .Context = methodContext
        }

            tokens.Add(tok)
            prevToken = tok
            prevOp = op
        Next

        Return tokens
    End Function



End Module
