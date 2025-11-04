'PROJECT NAME: VBA_TOKENIZER
'FILE DESCRIPTION: Modul pentru serializarea și deserializarea cache-urilor de proiecte VBA (multiple cache files)
'PATH: VBA_TOKENIZER/modSerializer.vb
Imports System.Collections.Concurrent
Imports System.IO
Imports System.IO.Compression
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Threading
Imports VBA_CORE.Logger
Imports VBA_CORE.modTypes
Imports VBA_CORE

Public Module modSerializer
    ' ============================================================
    '  CACHE FILE PATHS
    ' ============================================================
    Private ReadOnly CACHE_DIR As String = modGlobals.Global_ExportDir
    Private ReadOnly CACHE_MODULES As String = Path.Combine(CACHE_DIR, "cache_modules.dat")
    Private ReadOnly CACHE_FORMS As String = Path.Combine(CACHE_DIR, "cache_forms.dat")
    Private ReadOnly CACHE_RELATIONS As String = Path.Combine(CACHE_DIR, "cache_relations.dat")

    ' ============================================================
    '  CACHE CLASSES
    ' ============================================================

    <Serializable>
    Public Class ModulesCache
        Public Property Modules As ConcurrentDictionary(Of String, ModuleContainer)
        Public Property Timestamp As DateTime
        Public Property SourceHash As String
    End Class

    <Serializable>
    Public Class FormsReportsCache
        Public Property FormsReports As ConcurrentDictionary(Of String, FormReportContainer)
        Public Property Timestamp As DateTime
        Public Property SourceHash As String
    End Class


    ' ============================================================
    '  UNIFIED CACHE CLASS (Legacy support)
    ' ============================================================

    <Serializable>
    Public Class ProjectCache
        Public Property Modules As ConcurrentDictionary(Of String, ModuleContainer)
        Public Property FormsReports As ConcurrentDictionary(Of String, FormReportContainer)
        Public Property Timestamp As DateTime
        Public Property SourceHash As String
    End Class

    ' ============================================================
    '  SAVE FUNCTIONS
    ' ============================================================

    ''' <summary>
    ''' Salvează TOATE cache-urile (modules, forms, relations)
    ''' </summary>
    Public Function SaveAllCaches() As Boolean
        Try
            LogInfo("💾 Saving all caches...", 1)

            Dim results As New List(Of Boolean) From {
                SaveModulesCache(),
                SaveFormsReportsCache()
            }

            Dim success = results.All(Function(r) r)
            If success Then
                LogInfo("✅ All caches saved successfully!", 1)
            Else
                LogInfo("⚠️ Some caches failed to save", 1)
            End If

            Return success

        Catch ex As Exception
            LogError("SaveAllCaches", ex)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Salvează cache pentru modules
    ''' </summary>
    Public Function SaveModulesCache() As Boolean
        Try
            LogInfo("💾 Saving modules cache...", 2)

            Dim cache As New ModulesCache With {
                .Modules = GlobalModules,
                .Timestamp = DateTime.Now,
                .SourceHash = GetSourceHash()
            }

            Dim success = SerializeCache(cache, CACHE_MODULES)

            If success Then
                Dim size = New FileInfo(CACHE_MODULES).Length / 1024.0 / 1024.0
                LogInfo($"✅ Modules cache saved ({size:N2} MB)", 2)
            End If

            Return success

        Catch ex As Exception
            LogError("SaveModulesCache", ex)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Salvează cache pentru forms/reports
    ''' </summary>
    Public Function SaveFormsReportsCache() As Boolean
        Try
            LogInfo("💾 Saving forms/reports cache...", 2)

            Dim cache As New FormsReportsCache With {
                .FormsReports = GlobalFormsReports,
                .Timestamp = DateTime.Now,
                .SourceHash = GetSourceHash()
            }

            Dim success = SerializeCache(cache, CACHE_FORMS)

            If success Then
                Dim size = New FileInfo(CACHE_FORMS).Length / 1024.0 / 1024.0
                LogInfo($"✅ Forms/Reports cache saved ({size:N2} MB)", 2)
            End If

            Return success

        Catch ex As Exception
            LogError("SaveFormsReportsCache", ex)
            Return False
        End Try
    End Function

    ' ============================================================
    '  LOAD FUNCTIONS
    ' ============================================================

    ''' <summary>
    ''' Încarcă TOATE cache-urile (modules, forms, relations)
    ''' </summary>
    Public Function LoadAllCaches() As Boolean
        Try
            LogInfo("📂 Loading all caches...", 1)

            Dim modulesOk = LoadModulesCache()
            Dim formsOk = LoadFormsReportsCache()

            If modulesOk AndAlso formsOk Then
                LogInfo("✅ All caches loaded successfully!", 1)
                Return True
            Else
                LogInfo("❌ Critical caches failed to load", 1)
                Return False
            End If

        Catch ex As Exception
            LogError("LoadAllCaches", ex)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Încarcă cache pentru modules
    ''' </summary>
    Public Function LoadModulesCache() As Boolean
        Try
            If Not File.Exists(CACHE_MODULES) Then
                LogInfo("❌ No modules cache found.", 2)
                Return False
            End If

            LogInfo("📂 Loading modules cache...", 2)

            Dim cache As ModulesCache = DeserializeCache(Of ModulesCache)(CACHE_MODULES)

            If cache Is Nothing Then Return False

            LogInfo($"   → Timestamp: {cache.Timestamp:yyyy-MM-dd HH:mm:ss}", 2)

            If cache.SourceHash <> GetSourceHash() Then
                LogInfo("⚠️ Modules cache is invalid (source changed)", 2)
                Return False
            End If

            GlobalModules = cache.Modules
            LogInfo($"✅ Loaded {cache.Modules.Count} modules from cache", 2)

            Return True

        Catch ex As Exception
            LogError("LoadModulesCache", ex)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Încarcă cache pentru forms/reports
    ''' </summary>
    Public Function LoadFormsReportsCache() As Boolean
        Try
            If Not File.Exists(CACHE_FORMS) Then
                LogInfo("❌ No forms/reports cache found.", 2)
                Return False
            End If

            LogInfo("📂 Loading forms/reports cache...", 2)

            Dim cache As FormsReportsCache = DeserializeCache(Of FormsReportsCache)(CACHE_FORMS)

            If cache Is Nothing Then Return False

            LogInfo($"   → Timestamp: {cache.Timestamp:yyyy-MM-dd HH:mm:ss}", 2)

            If cache.SourceHash <> GetSourceHash() Then
                LogInfo("⚠️ Forms/Reports cache is invalid (source changed)", 2)
                Return False
            End If

            GlobalFormsReports = cache.FormsReports
            LogInfo($"✅ Loaded {cache.FormsReports.Count} forms/reports from cache", 2)

            Return True

        Catch ex As Exception
            LogError("LoadFormsReportsCache", ex)
            Return False
        End Try
    End Function

    ' ============================================================
    '  LEGACY SUPPORT (Single cache file)
    ' ============================================================

    ''' <summary>
    ''' Salvează cache unificat (backward compatibility)
    ''' </summary>
    Public Function SaveCache() As Boolean
        Try
            LogInfo("💾 Saving unified cache (legacy)...", 1)

            Dim cachePath = Path.Combine(CACHE_DIR, "cache.dat")

            Dim cache As New ProjectCache With {
                .Modules = GlobalModules,
                .FormsReports = GlobalFormsReports,
                .Timestamp = DateTime.Now,
                .SourceHash = GetSourceHash()
            }

            Dim success = SerializeCache(cache, cachePath)

            If success Then
                Dim size = New FileInfo(cachePath).Length / 1024.0 / 1024.0
                LogInfo($"✅ Unified cache saved ({size:N2} MB)", 1)
            End If

            Return success

        Catch ex As Exception
            LogError("SaveCache", ex)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Încarcă cache unificat (backward compatibility)
    ''' </summary>
    Public Function LoadCache() As Boolean
        Try
            Dim cachePath = Path.Combine(CACHE_DIR, "cache.dat")

            If Not File.Exists(cachePath) Then
                LogInfo("❌ No unified cache found.", 1)
                Return False
            End If

            LogInfo("📂 Loading unified cache (legacy)...", 1)

            Dim cache As ProjectCache = DeserializeCache(Of ProjectCache)(cachePath)

            If cache Is Nothing Then Return False

            LogInfo($"   → Extracted: {cache.Modules.Count} modules, {cache.FormsReports.Count} forms/reports", 1)
            LogInfo($"   → Timestamp: {cache.Timestamp:yyyy-MM-dd HH:mm:ss}", 1)

            If cache.SourceHash <> GetSourceHash() Then
                LogInfo("⚠️ Cache is invalid (source changed)", 1)
                Return False
            End If

            GlobalModules = cache.Modules
            GlobalFormsReports = cache.FormsReports

            LogInfo($"✅ Unified cache loaded successfully", 1)

            Return True

        Catch ex As Exception
            LogError("LoadCache", ex)
            Return False
        End Try
    End Function

    ' ============================================================
    '  HELPER FUNCTIONS
    ' ============================================================

    ''' <summary>
    ''' Serializează un obiect în fișier cu compresie GZip
    ''' </summary>
    Private Function SerializeCache(Of T)(cache As T, filePath As String) As Boolean
        Try
            Using fs As New FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None)
                Using gz As New GZipStream(fs, CompressionLevel.Optimal)
                    Dim bf As New BinaryFormatter() With {
                        .AssemblyFormat = Runtime.Serialization.Formatters.FormatterAssemblyStyle.Simple
                    }
                    bf.Serialize(gz, cache)
                End Using
            End Using

            Return True

        Catch ex As Exception
            LogError($"SerializeCache({Path.GetFileName(filePath)})", ex)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Deserializează un obiect din fișier cu animație progress
    ''' </summary>
    Private Function DeserializeCache(Of T As Class)(filePath As String) As T
        Try
            Dim cache As T = Nothing
            Dim cancelToken As Boolean = False
            Dim spinnerChars As Char() = {"|"c, "/"c, "-"c, "\"c}
            Dim spinIndex As Integer = 0
            Dim lastPct As Integer = -1
            Dim isConsole As Boolean = Environment.UserInteractive AndAlso Not Console.IsOutputRedirected
            Dim fileName As String = Path.GetFileName(filePath)

            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read)
                Using gz As New GZipStream(fs, CompressionMode.Decompress)
                    Dim bf As New BinaryFormatter()

                    ' Thread pentru progress animation
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
                                                                            Dim spinner As Char = spinnerChars(spinIndex Mod spinnerChars.Length)
                                                                            spinIndex += 1
                                                                            Console.Write(vbCr & $"   {spinner} Deserializing {fileName}: {curPct,3}% complete   ")
                                                                        Else
                                                                            dots = (dots + 1) Mod 4
                                                                            Dim anim = New String("."c, dots)
                                                                            LogInfo($"   Deserializing {fileName}: {curPct,3}% {anim}", 3, True)
                                                                        End If
                                                                    End If
                                                                Catch
                                                                End Try
                                                                Thread.Sleep(120)
                                                            End While

                                                            If isConsole Then
                                                                Console.Write(vbCr & $"   ✅ {fileName} deserialized!                     " & vbCrLf)
                                                            End If
                                                        End Sub)

                    cache = CType(bf.Deserialize(gz), T)

                    cancelToken = True
                    progressTask.Wait()
                End Using
            End Using

            Return cache

        Catch ex As Exception
            LogError($"DeserializeCache({Path.GetFileName(filePath)})", ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Calculează hash pentru validarea cache-ului
    ''' </summary>
    Private Function GetSourceHash() As String
        Try
            Dim fi As New FileInfo(modGlobals.Global_DBPath)
            Return $"{fi.Length}_{fi.LastWriteTimeUtc.Ticks}"
        Catch
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Șterge toate cache-urile
    ''' </summary>
    Public Sub ClearAllCaches()
        Try
            LogInfo("🗑️ Clearing all caches...", 1)

            If File.Exists(CACHE_MODULES) Then
                File.Delete(CACHE_MODULES)
                LogInfo("   → Modules cache deleted", 2)
            End If

            If File.Exists(CACHE_FORMS) Then
                File.Delete(CACHE_FORMS)
                LogInfo("   → Forms/Reports cache deleted", 2)
            End If

            If File.Exists(CACHE_RELATIONS) Then
                File.Delete(CACHE_RELATIONS)
                LogInfo("   → Relations cache deleted", 2)
            End If

            Dim legacyCache = Path.Combine(CACHE_DIR, "cache.dat")
            If File.Exists(legacyCache) Then
                File.Delete(legacyCache)
                LogInfo("   → Legacy cache deleted", 2)
            End If

            LogInfo("✅ All caches cleared", 1)

        Catch ex As Exception
            LogError("ClearAllCaches", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Verifică dacă există cache-uri valide
    ''' </summary>
    Public Function HasValidCaches() As (Modules As Boolean, Forms As Boolean, Relations As Boolean)
        Try
            Dim sourceHash = GetSourceHash()

            Dim modulesValid As Boolean = False
            Dim formsValid As Boolean = False
            Dim relationsValid As Boolean = False

            ' Check modules
            If File.Exists(CACHE_MODULES) Then
                Try
                    Dim cache = DeserializeCache(Of ModulesCache)(CACHE_MODULES)
                    modulesValid = (cache IsNot Nothing AndAlso cache.SourceHash = sourceHash)
                Catch
                End Try
            End If

            ' Check forms
            If File.Exists(CACHE_FORMS) Then
                Try
                    Dim cache = DeserializeCache(Of FormsReportsCache)(CACHE_FORMS)
                    formsValid = (cache IsNot Nothing AndAlso cache.SourceHash = sourceHash)
                Catch
                End Try
            End If

            Return (modulesValid, formsValid, relationsValid)

        Catch ex As Exception
            LogError("HasValidCaches", ex)
            Return (False, False, False)
        End Try
    End Function

End Module
