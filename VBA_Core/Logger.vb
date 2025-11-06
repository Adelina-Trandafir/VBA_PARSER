Imports System.IO
Imports System.Text
Imports System.Collections.Concurrent
Imports System.Threading

''' <summary>
''' Logger asincron, thread-safe, fără output în consolă.
''' Scrie în fișiere + trimite evenimente UI (WPF) fără să blocheze firele de lucru.
''' </summary>
Public Module Logger
    ' ============================================================
    ' 🔹 EVENIMENT pentru UI (ascultat de WPF)
    ' ============================================================
    Public Event OnLogMessage(ByVal message As String, ByVal level As Integer)
    Public UI_InvokeSafe As Action(Of String, Integer)

    ' ============================================================
    ' 🔹 STRUCTURI INTERNE
    ' ============================================================
    Private ReadOnly _queue As New BlockingCollection(Of (msg As String, level As Integer, file As String))
    Private _workerThread As Thread
    Private _isRunning As Boolean = False

    Private ReadOnly ERR_LOG As String = Path.Combine(modGlobals.Global_ExportDir, "Errors.log")
    Private ReadOnly INFO_LOG As String = Path.Combine(modGlobals.Global_ExportDir, "Run.log")

    Private swError As StreamWriter
    Private swInfo As StreamWriter
    Private messageCounter As Integer = 0
    Private Const FLUSH_INTERVAL As Integer = 10

    ' ============================================================
    ' 🔹 INITIALIZARE / CLEANUP
    ' ============================================================
    Public Sub InitLogger()
        If _isRunning Then Exit Sub
        Try
            Directory.CreateDirectory(modGlobals.Global_ExportDir)

            swError = New StreamWriter(ERR_LOG, append:=True, encoding:=Encoding.UTF8) With {.AutoFlush = False}
            swInfo = New StreamWriter(INFO_LOG, append:=True, encoding:=Encoding.UTF8) With {.AutoFlush = False}

            _isRunning = True
            _workerThread = New Thread(AddressOf LogWorker) With {
            .IsBackground = True,
            .Name = "LoggerThread"
        }
            _workerThread.Start()

            ' ✅ înlocuiește linia veche cu:
            EnsureShutdownHandlers()

        Catch ex As Exception
            File.AppendAllText("LoggerInit_Fail.txt", ex.ToString() & Environment.NewLine)
        End Try
    End Sub


    Private Sub LogWorker()
        Try
            For Each entry In _queue.GetConsumingEnumerable()
                Dim target = If(entry.file = "ERR", swError, swInfo)
                target.WriteLine(entry.msg)
                messageCounter += 1
                If messageCounter Mod FLUSH_INTERVAL = 0 Then
                    target.Flush()
                End If
            Next
        Catch ex As Exception
            File.AppendAllText("LoggerWorker_Fail.txt", ex.ToString() & Environment.NewLine)
        End Try
    End Sub

    Public Sub CloseLogger(sender As Object, e As EventArgs)
        Try
            _isRunning = False
            _queue.CompleteAdding()

            If _workerThread IsNot Nothing AndAlso _workerThread.IsAlive Then
                _workerThread.Join(2000)
            End If

            swError?.Flush()
            swInfo?.Flush()
            swError?.Close()
            swInfo?.Close()
        Catch ex As Exception
            File.AppendAllText("LoggerClose_Fail.txt", ex.ToString() & Environment.NewLine)
        End Try
    End Sub

    ' ============================================================
    ' ❌ ERROR LOGGING
    ' ============================================================
    ''' <summary>
    ''' Scrie o eroare în log (întotdeauna în fișier și în UI).
    ''' ExtendedVerbose afectează doar nivelul intern, nu afișarea.
    ''' </summary>
    ''' <param name="context">Contextul sau sursa erorii (ex: numele funcției).</param>
    ''' <param name="ex">Obiectul excepției.</param>
    ''' <param name="verbosityLevel">Nivelul de detaliu al erorii (1 = important, 2 = verbose, 3 = extended).</param>
    ''' <param name="force">Dacă True, se scrie mereu, indiferent de nivelul global.</param>
    Public Sub LogError(context As String,
                    ex As Exception,
                    Optional verbosityLevel As Integer = 1,
                    Optional force As Boolean = True)

        Try
            ' === Inițializare dacă nu e pornit ===
            If Not _isRunning Then InitLogger()

            ' === Filtrare globală (doar pentru oprirea logării, nu pentru UI) ===
            If Not force AndAlso (verbosityLevel > modGlobals.Global_VerboseLevel) Then Exit Sub

            ' === Stack trace friendly info ===
            Dim trace As New System.Diagnostics.StackTrace(ex, True)
            Dim frame As System.Diagnostics.StackFrame =
            trace.GetFrames()?.FirstOrDefault(Function(f)
                                                  Dim ns = f.GetMethod()?.DeclaringType?.Namespace
                                                  Return ns IsNot Nothing AndAlso
                                                         Not ns.StartsWith("System") AndAlso
                                                         Not ns.StartsWith("Microsoft")
                                              End Function)

            Dim methodName As String = If(frame?.GetMethod()?.Name, "?")
            Dim fileName As String = If(Path.GetFileName(frame?.GetFileName()), "?")
            Dim lineNum As Integer = If(frame?.GetFileLineNumber(), 0)

            ' === Construiește linia de log pentru fișier ===
            Dim msg As String =
            $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ❌ {context}: {ex.GetType().Name} - {ex.Message}" & Environment.NewLine &
            $"   at {methodName} ({fileName}:{lineNum})"

            ' 🔹 Adaugă în coada de scriere (merge oricum în fișier)
            _queue.Add((msg, verbosityLevel, "ERROR"))

            ' 🔹 Trimite întotdeauna și în UI
            If UI_InvokeSafe IsNot Nothing Then
                UI_InvokeSafe.Invoke(msg, 2)
            Else
                RaiseEvent OnLogMessage(msg, 2)
            End If


            ' === InnerException (dacă există) ===
            If ex.InnerException IsNot Nothing Then
                Dim innerMsg = $"   Inner: {ex.InnerException.Message}"
                _queue.Add((innerMsg, verbosityLevel, "ERROR"))
                RaiseEvent OnLogMessage(innerMsg, 2)
            End If

        Catch failEx As Exception
            ' fallback absolut (în caz că threadul logger e blocat)
            Try
                File.AppendAllText("LoggerError_Fail.txt",
                $"[{DateTime.Now:HH:mm:ss}] {context}: {ex.Message}{Environment.NewLine}{failEx}{Environment.NewLine}",
                Encoding.UTF8)
            Catch
            End Try
        End Try
    End Sub


    ' ============================================================
    ' ℹ️ INFO LOGGING
    ' ============================================================
    ''' <summary>
    ''' Scrie un mesaj informativ în log (UI + fișier), cu nivel de detaliu controlabil.
    ''' </summary>
    ''' <param name="msg">Mesajul text.</param>
    ''' <param name="verbosityLevel">
    ''' 0 = tăcere totală,
    ''' 1 = normal (doar esențial),
    ''' 2 = verbose (afișează tot),
    ''' 3 = extended verbose (doar fișier, fără UI).
    ''' </param>
    ''' <param name="force">Dacă True, mesajul se scrie indiferent de nivelul global.</param>
    Public Sub LogInfo(msg As String, Optional verbosityLevel As Integer = 1, Optional force As Boolean = False)
        Try
            ' === Inițializare dacă nu e pornit ===
            If Not _isRunning Then InitLogger()

            ' === Filtrare globală ===
            If Not force AndAlso (verbosityLevel > modGlobals.Global_VerboseLevel) Then Exit Sub

            ' === Construiește linia de log pentru fișier ===
            Dim line As String = $"[{DateTime.Now:HH:mm:ss}] {msg}"
            _queue.Add((line, verbosityLevel, "INFO"))

            ' === Trimite și către UI (RichTextBox) dacă e cazul ===
            '  - Normal & Verbose => mereu în UI
            '  - ExtendedVerbose => doar dacă utilizatorul e în mod verbose, dar cu marcaj [EXT]
            If verbosityLevel <= 2 Then
                If UI_InvokeSafe IsNot Nothing Then
                    UI_InvokeSafe.Invoke(msg, 2)
                Else
                    RaiseEvent OnLogMessage(msg, 2)
                End If
            End If

            If ExtendedVerbose Then RaiseEvent OnLogMessage("[EXT] " & msg, 2)

        Catch ex As Exception
            ' fallback sigur (doar pentru debugging)
            Try
                File.AppendAllText("LoggerInfo_Fail.txt", ex.ToString() & Environment.NewLine)
            Catch
            End Try
        End Try
    End Sub


    ''' <summary>
    ''' Închide loggerul și eliberează fișierele în siguranță.
    ''' Se apelează la închiderea aplicației.
    ''' </summary>
    Public Sub ShutdownLogger()
        Try
            If swInfo IsNot Nothing Then
                swInfo.Flush()
                swInfo.Close()
                swInfo.Dispose()
                swInfo = Nothing
            End If

            If swError IsNot Nothing Then
                swError.Flush()
                swError.Close()
                swError.Dispose()
                swError = Nothing
            End If

        Catch ex As Exception
            ' Nu facem nimic aici — oricum aplicația se închide
        End Try
    End Sub

    ' =====================================================
    ' 🧩 AUTOMATIC SAFE SHUTDOWN HANDLERS
    ' =====================================================
    Private initializedHandlers As Boolean = False

    ''' <summary>
    ''' Atașează automat evenimente pentru a închide loggerul
    ''' în caz de ieșire, excepție globală sau terminare bruscă.
    ''' Se apelează o singură dată, la prima utilizare.
    ''' </summary>
    Private Sub EnsureShutdownHandlers()
        If initializedHandlers Then Exit Sub
        initializedHandlers = True

        Try
            AddHandler AppDomain.CurrentDomain.ProcessExit, AddressOf AutoShutdown
            AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf AutoShutdownUnhandled
        Catch
            ' Ignoră – unele contexte (ex: test runner) nu permit adăugarea
        End Try
    End Sub

    Private Sub AutoShutdown(sender As Object, e As EventArgs)
        ShutdownLogger()
    End Sub

    Private Sub AutoShutdownUnhandled(sender As Object, e As UnhandledExceptionEventArgs)
        Try
            File.AppendAllText("Logger_Unhandled.txt",
            $"[{DateTime.Now:HH:mm:ss}] Unhandled: {e.ExceptionObject}{Environment.NewLine}",
            Encoding.UTF8)
        Catch
        End Try

        ShutdownLogger()
    End Sub

End Module
