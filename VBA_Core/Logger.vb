Imports System.IO
Imports System.Text
Imports System.Threading
Imports VBA_CORE
Imports VBA_CORE.VBA_CORE   ' <-- accesăm direct variabilele globale

Public Module Logger
    ' ===============================
    ' 🧭 EVENTS pentru UI
    ' ===============================
    Public Event OnLogMessage(ByVal message As String, ByVal level As Integer)

    ' ===============================
    ' 🧰 CONFIG RUNTIME
    ' ===============================
    Public mainThreadId As Integer = Thread.CurrentThread.ManagedThreadId

    ' ===============================
    ' 🗃️ FILE PATHS
    ' ===============================
    Private ReadOnly ERR_LOG As String = Path.Combine(modGlobals.Global_ExportDir, "Errors.log")
    Private ReadOnly INFO_LOG As String = Path.Combine(modGlobals.Global_ExportDir, "Run.log")

    ' ===============================
    ' ❌ ERROR LOGGING
    ' ===============================
    Public Sub LogError(context As String, ex As Exception)
        Try
            Directory.CreateDirectory(modGlobals.Global_ExportDir)

            ' extrage primul frame util din stack
            Dim trace As New System.Diagnostics.StackTrace(ex, True)
            Dim frame As System.Diagnostics.StackFrame = Nothing
            Dim methodName As String = "?"
            Dim lineNum As Integer = 0
            Dim fileName As String = "?"

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

            RaiseEvent OnLogMessage("❌ " & msg, 2)

            If ex.InnerException IsNot Nothing Then
                File.AppendAllText(ERR_LOG, "   Inner: " & ex.InnerException.Message & Environment.NewLine, Encoding.UTF8)
            End If

        Catch
            Console.WriteLine($"[E] {context}: {ex.Message}")
        End Try
    End Sub

    ' ===============================
    ' ℹ️ INFO LOGGING
    ' ===============================
    Public Sub LogInfo(msg As String, Optional force As Boolean = False)
        Static firstWrite As Boolean = True
        Const RUN_LOG As String = "C:\exportvba\runlog.txt"

        If Global_VerboseLevel = 0 AndAlso Not force Then Exit Sub

        Try
            SyncLock Console.Out
                Console.WriteLine(msg)
            End SyncLock

            RaiseEvent OnLogMessage(msg, Global_VerboseLevel)

            ' log normal în fișier principal
            If Global_VerboseLevel > 1 Then
                Directory.CreateDirectory(Global_ExportDir)
                File.AppendAllText(INFO_LOG, $"[{DateTime.Now:HH:mm:ss}] {msg}{Environment.NewLine}", Encoding.UTF8)
            End If

            ' log separat în C:\exportvba\runlog.txt
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

    ' ===============================
    ' 💬 SAFE CONSOLE
    ' ===============================
    Public Sub SafeConsoleWriteLine(msg As String)
        If Global_VerboseLevel = 0 Then Exit Sub
        If Thread.CurrentThread.ManagedThreadId <> mainThreadId Then Exit Sub
        SyncLock Console.Out
            Console.WriteLine(msg)
        End SyncLock
        RaiseEvent OnLogMessage(msg, Global_VerboseLevel)
    End Sub
End Module
