Imports System.IO
Imports System.Text
Imports System.Threading

Public Module Logger
    ' ===============================
    ' 🔧 CONFIG
    ' ===============================
    Public ReadOnly DB_PATH As String = "C:\AVACONT\Avacont.accdb"
    Public ReadOnly EXPORT_DIR As String = "C:\ExportVBA"

    Public useParallel As Boolean = True            ' --nothreads => False
    Public verboseLevel As Integer = 1              ' 0=quiet, 1=normal, 2=verbose
    Public mainThreadId As Integer = Thread.CurrentThread.ManagedThreadId

    ' ===============================
    ' 🗃️ RAM caches (global vizibile)
    ' ===============================
    'Public fileCache As New Dictionary(Of String, PreloadedFile)(StringComparer.OrdinalIgnoreCase)

    'Public mapCache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

    ' ===============================
    ' 🧾 LOGGING
    ' ===============================
    Private ReadOnly ERR_LOG As String = Path.Combine(EXPORT_DIR, "Errors.log")
    Private ReadOnly INFO_LOG As String = Path.Combine(EXPORT_DIR, "Run.log")

    Public Sub LogError(context As String, ex As Exception)
        Try
            Directory.CreateDirectory(EXPORT_DIR)

            ' === extrage primul frame util din stack ===
            Dim trace As New System.Diagnostics.StackTrace(ex, True)
            Dim frame As System.Diagnostics.StackFrame = Nothing
            Dim methodName As String = "?"
            Dim lineNum As Integer = 0
            Dim fileName As String = "?"

            ' caută primul frame din codul tău (nu din mscorlib)
            For Each f In trace.GetFrames()
                If f.GetMethod()?.DeclaringType Is Nothing Then Continue For
                Dim ns = f.GetMethod().DeclaringType.Namespace
                If ns Is Nothing OrElse ns.StartsWith("System") OrElse ns.StartsWith("Microsoft") Then Continue For
                frame = f
                Exit For
            Next

            If frame IsNot Nothing Then
                methodName = frame.GetMethod().DeclaringType.FullName & "." & frame.GetMethod().Name
                lineNum = frame.GetFileLineNumber()
                fileName = Path.GetFileName(frame.GetFileName())
            End If

            Dim msg As String =
            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {context}: {ex.GetType().Name} - {ex.Message}" & Environment.NewLine &
            $"   at {methodName} ({fileName}:{lineNum})"

            ' scrie în fișier + consolă
            File.AppendAllText(ERR_LOG, msg & Environment.NewLine, Encoding.UTF8)
            SyncLock Console.Out
                Console.ForegroundColor = ConsoleColor.Red
                Console.WriteLine("❌ " & msg)
                Console.ResetColor()
            End SyncLock

            ' scrie și inner exception, dacă există
            If ex.InnerException IsNot Nothing Then
                File.AppendAllText(ERR_LOG, "   Inner: " & ex.InnerException.Message & Environment.NewLine, Encoding.UTF8)
            End If

        Catch
            ' fallback minimal (evită excepții în logger)
            Console.WriteLine($"[E] {context}: {ex.Message}")
        End Try
    End Sub

    Public Sub LogInfo(msg As String, Optional force As Boolean = False)
        Static firstWrite As Boolean = True
        Const RUN_LOG As String = "C:\exportvba\runlog.txt"

        If verboseLevel = 0 AndAlso Not force Then Exit Sub

        Try
            SyncLock Console.Out
                Console.WriteLine(msg)
            End SyncLock

            ' === log normal în fișier principal ===
            If verboseLevel > 1 Then
                Directory.CreateDirectory(EXPORT_DIR)
                File.AppendAllText(INFO_LOG, $"[{DateTime.Now:HH:mm:ss}] {msg}{Environment.NewLine}", Encoding.UTF8)
            End If

            ' === log separat în C:\exportvba\runlog.txt ===
            Directory.CreateDirectory("C:\exportvba")

            If firstWrite Then
                File.WriteAllText(RUN_LOG, $"[{DateTime.Now:HH:mm:ss}] {msg}{Environment.NewLine}", Encoding.UTF8)
                firstWrite = False
            Else
                File.AppendAllText(RUN_LOG, $"[{DateTime.Now:HH:mm:ss}] {msg}{Environment.NewLine}", Encoding.UTF8)
            End If

        Catch
        End Try
    End Sub

    Public Sub SafeConsoleWriteLine(msg As String)
        If verboseLevel = 0 Then Exit Sub
        If Thread.CurrentThread.ManagedThreadId <> mainThreadId Then Exit Sub
        SyncLock Console.Out
            Console.WriteLine(msg)
        End SyncLock
    End Sub

End Module
