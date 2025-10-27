Imports System.IO
Imports System.IO.Compression
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Threading
Imports VBA_CORE.Logger
Imports VBA_CORE.modTypes
Imports VBA_CORE.VBA_CORE

Public Module modSerializer
    Private ReadOnly CACHE_FILE As String = Path.Combine(modGlobals.Global_ExportDir, "cache.dat")

    <Serializable>
    Public Class ProjectCache
        Public Property Modules As Dictionary(Of String, ModuleContainer)
        Public Property FormsReports As Dictionary(Of String, FormReportContainer)
        Public Property Timestamp As DateTime
        Public Property SourceHash As String
    End Class

    Public Function SaveCache() As Boolean
        Try
            LogInfo("💾 Salvez cache binar...")

            Dim cache As New ProjectCache With {
                .Modules = GlobalModules,
                .FormsReports = GlobalFormsReports,
                .Timestamp = DateTime.Now,
                .SourceHash = GetSourceHash()
            }

            Using fs As New FileStream(CACHE_FILE, FileMode.Create, FileAccess.Write, FileShare.None)
                Using gz As New GZipStream(fs, CompressionLevel.Optimal)
                    Dim bf As New BinaryFormatter() With {
                        .AssemblyFormat = Runtime.Serialization.Formatters.FormatterAssemblyStyle.Simple
                    }
#Disable Warning SYSLIB0011
                    bf.Serialize(gz, cache)
#Enable Warning SYSLIB0011
                End Using
            End Using

            Dim size = New FileInfo(CACHE_FILE).Length / 1024.0 / 1024.0
            LogInfo($"✅ Cache binar salvat ({size:N2} MB)")
            Return True

        Catch ex As Exception
            LogError("SaveCache", ex)
            Return False
        End Try
    End Function

    Public Function LoadCache() As Boolean
        Try
            If Not File.Exists(CACHE_FILE) Then
                LogInfo("❌ Nu există cache salvat.")
                Return False
            End If

            LogInfo("📂 Încarc cache binar...")

            Dim cache As ProjectCache
            Dim cancelToken As Boolean = False
            Dim spinnerChars As Char() = {"|", "/", "-", "\"}
            Dim spinIndex As Integer = 0
            Dim lastPct As Integer = -1
            Dim isConsole As Boolean = Environment.UserInteractive AndAlso Not Console.IsOutputRedirected

            Using fs As New FileStream(CACHE_FILE, FileMode.Open, FileAccess.Read, FileShare.Read)
                Using gz As New GZipStream(fs, CompressionMode.Decompress)
                    Dim bf As New BinaryFormatter()

                    ' === Thread separat pentru progres ===
                    Dim progressTask As Task = Task.Run(Sub()
                                                            Dim dots As Integer = 0
                                                            While Not cancelToken
                                                                Try
                                                                    Dim pct As Double = 0
                                                                    If fs.Length > 0 Then
                                                                        pct = (fs.Position / fs.Length) * 100
                                                                    End If

                                                                    Dim curPct As Integer = CInt(Math.Min(100, Math.Round(pct)))
                                                                    If curPct <> lastPct Then
                                                                        lastPct = curPct

                                                                        If isConsole Then
                                                                            ' Consolă reală: animăm în aceeași linie
                                                                            Dim spinner As Char = spinnerChars(spinIndex Mod spinnerChars.Length)
                                                                            spinIndex += 1
                                                                            Console.Write(vbCr & $"   {spinner} Deserializare cache: {curPct,3}% complet   ")
                                                                        Else
                                                                            ' Fallback pentru VS Output / PowerShell
                                                                            dots = (dots + 1) Mod 4
                                                                            Dim anim = New String("."c, dots)
                                                                            LogInfo($"   Deserializare cache{anim} {curPct,3}% complet")
                                                                        End If
                                                                    End If
                                                                Catch
                                                                End Try
                                                                Thread.Sleep(120)
                                                            End While

                                                            ' finalizează linia dacă e consolă
                                                            If isConsole Then
                                                                Console.Write(vbCr & "   ✅ Deserializare completă!                     " & vbCrLf)
                                                            End If
                                                        End Sub)

#Disable Warning SYSLIB0011
                    cache = CType(bf.Deserialize(gz), ProjectCache)
#Enable Warning SYSLIB0011

                    cancelToken = True
                    progressTask.Wait()
                End Using
            End Using

            LogInfo($"   → Cache extras: {cache.Modules.Count} module, {cache.FormsReports.Count} forms/reports")
            LogInfo($"   → Timestamp cache: {cache.Timestamp:yyyy-MM-dd HH:mm:ss}")

            If cache.SourceHash <> GetSourceHash() Then
                LogInfo("⚠️ Sursele s-au modificat. Cache invalid.")
                Return False
            End If

            ' === Populate globale ===
            LogInfo("🔄 Procesez module din cache...")
            GlobalModules = cache.Modules
            Dim idx As Integer = 0
            For Each moduleEntry In cache.Modules
                idx += 1
                If idx Mod 50 = 0 Then LogInfo($"   → {idx}/{cache.Modules.Count} module procesate...")
            Next

            LogInfo("🔄 Procesez forms/reports din cache...")
            GlobalFormsReports = cache.FormsReports
            idx = 0
            For Each formEntry In cache.FormsReports
                idx += 1
                If idx Mod 20 = 0 Then LogInfo($"   → {idx}/{cache.FormsReports.Count} forms/reports procesate...")
            Next

            LogInfo($"✅ Cache încărcat complet ({cache.Modules.Count} module, {cache.FormsReports.Count} forms/reports)")
            Return True

        Catch ex As Exception
            LogError("LoadCache", ex)
            Return False
        End Try
    End Function

    Private Function GetSourceHash() As String
        Try
            Dim fi As New FileInfo(modGlobals.Global_DBPath)
            Return $"{fi.Length}_{fi.LastWriteTimeUtc.Ticks}"
        Catch
            Return ""
        End Try
    End Function
End Module
