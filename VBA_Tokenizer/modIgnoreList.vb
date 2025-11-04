'PROJECT NAME: VBA_TOKENIZER
'FILE DESCRIPTION: Module pentru lista de ignorare a tokenilor VBA
'PATH: VBA_TOKENIZER/modIgnoreList.vb

Imports System.Text.RegularExpressions
Imports VBA_CORE.modTypes

Public Module VBAIgnoreList

    ' === prefixuri de obiecte globale (nu sunt module definite de tine) ===
    Private ReadOnly IgnoredPrefixes As String() = {
        "VBA.", "Vb.", "Application.", "CurrentDb.", "DBEngine.", "DoCmd.",
        "Forms!", "Reports!", "Screen.", "Err.", "Debug.", "Attribute*",
        "Recordset.", "Form.", "Report."
    }

    ' === funcții / simboluri standard (fără prefix) ===
    Private ReadOnly IgnoredNames As HashSet(Of String) =
        New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "left", "right", "mid", "trim", "ltrim", "rtrim",
            "instr", "instrrev", "replace", "split", "join", "format", "chr", "asc", "val", "str", "len", "nz",
            "msgbox", "inputbox", "now", "date", "time", "timer", "weekday", "year", "month", "day",
            "cstr", "clng", "cint", "cdbl", "cdate", "csng", "ccur", "cbool",
            "isnull", "isempty", "isnumeric", "iserror", "not", "and", "or",
            "array", "ubound", "lbound", "err", "me", "nothing", "true", "false",
            "currentproject", "currentdb", "forms", "reports"
        }

    ' === proprietăți / controale tipice accesibile prin Me. ===
    Private ReadOnly IgnoredMeMembers As HashSet(Of String) =
        New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "name", "caption", "value", "text", "tag", "visible",
            "enabled", "locked", "controlsource", "recordsource",
            "controls", "section", "hasdata", "form", "report",
            "openargs", "requery", "sourceobject"
        }

    ' === proprietăți tipice Form/Report (de ignorat chiar și pe alt obiect) ===
    Private ReadOnly IgnoredFormProps As HashSet(Of String) =
        New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "controls", "recordsource", "rowsource", "controlsource",
            "sourceobject", "caption", "visible", "enabled", "locked",
            "hasdata", "section", "form", "report", "allowadditions",
            "allowdeletions", "allowedits", "requery", "openargs",
            "setfocus", "sellength", "selstart", "selend", "VB_VarHelpID"
        }

    ' === regex pentru identificarea operatorilor / constantelor VBA ===
    Private ReadOnly IgnorePatterns As Regex() = {
        New Regex("^\d+$"),               ' numere simple
        New Regex("^#.*#$"),              ' date #2025-10-22#
        New Regex("^'.*'$"),              ' stringuri
        New Regex("^"".*""$"),            ' literal string
        New Regex("^[\(\)\+\-\*/\\\^&]+$") ' operatori
    }

    ' =====================================================
    ' FUNCȚIE UNIFICATĂ DE DECIZIE IGNORARE (contextuală)
    ' =====================================================
    Public Function ShouldIgnore(token As String,
                                 Optional currentName As String = "",
                                 Optional currentType As String = "",
                                 Optional fci As FormReportContainer = Nothing) As Boolean

        If String.IsNullOrWhiteSpace(token) Then Return True
        Dim t As String = token.Trim()

        ' ===== normalizează context form/report + asigură fci =====
        Dim inFormOrReport As Boolean =
            currentType.Equals("Forms", StringComparison.OrdinalIgnoreCase) OrElse
            currentType.Equals("Reports", StringComparison.OrdinalIgnoreCase) OrElse
            currentType.Equals("Form", StringComparison.OrdinalIgnoreCase) OrElse
            currentType.Equals("Report", StringComparison.OrdinalIgnoreCase)

        If fci Is Nothing AndAlso inFormOrReport AndAlso Not String.IsNullOrEmpty(currentName) Then
            GlobalFormsReports.TryGetValue(currentName, fci)
        End If

        ' ===== 1) bloc: controale reale din formular (bare / Me. / Me!)
        If inFormOrReport AndAlso fci IsNot Nothing Then
            Dim isId As Boolean = Regex.IsMatch(t, "^[A-Za-z_]\w*(?:\.[A-Za-z_]\w*)?$")

            ' 1.a Me!X sau orice cu "!"
            If t.Contains("!") Then Return True

            ' 1.b Me.X  (X = control real sau proprietate standard)
            If t.StartsWith("Me.", StringComparison.OrdinalIgnoreCase) Then
                Dim member = t.Substring(3)
                If IgnoredMeMembers.Contains(member) OrElse IgnoredFormProps.Contains(member) Then Return True
                If fci.Controls.Any(Function(c) c.Name.Equals(member, StringComparison.OrdinalIgnoreCase)) Then Return True
            End If

            ' 1.c Bare: X sau X.Prop — dacă X e un control real, ignorăm
            If isId Then
                Dim parts = t.Split("."c)
                Dim root = parts(0)
                If fci.Controls.Any(Function(c) c.Name.Equals(root, StringComparison.OrdinalIgnoreCase)) Then
                    Return True
                End If
            End If
        End If

        ' ===== 2) prefixe standard / built-in etc. (logica existentă)
        For Each p In IgnoredPrefixes
            If t.StartsWith(p, StringComparison.OrdinalIgnoreCase) Then Return True
        Next

        For Each rg In IgnorePatterns
            If rg.IsMatch(t) Then Return True
        Next

        t = Regex.Replace(t, "[\(\)]", "")

        If IgnoredNames.Contains(t) Then Return True

        ' orice sufix proprietate tipică form/report (ex: X.Caption)
        Dim dotPos = t.LastIndexOf("."c)
        If dotPos > 0 Then
            Dim suffix = t.Substring(dotPos + 1)
            If IgnoredFormProps.Contains(suffix) Then Return True
        End If

        Return False
    End Function

End Module
