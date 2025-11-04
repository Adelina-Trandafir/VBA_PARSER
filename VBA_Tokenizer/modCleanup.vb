Imports VBA_CORE
Imports VBA_CORE.modTypes

Module modCleanup
    ''' <summary>
    ''' Elimină referințele circulare (WithBlockContext, SourceLine, WithBlockContext etc.)
    ''' din toate modulele (Forms/Reports/Class/Modules) din GlobalSources.
    ''' Se recomandă apelarea acestei funcții după terminarea tokenizării și rezoluției simbolice.
    ''' </summary>
    Public Sub CleanupWithReferences()
        Try
            If GlobalSources Is Nothing OrElse GlobalSources.Count = 0 Then Exit Sub

            Dim opts As New ParallelOptions With {.MaxDegreeOfParallelism = Environment.ProcessorCount}

            Dim runner As Action =
            Sub()
                If modGlobals.Global_UseParallel Then
                    Parallel.ForEach(GlobalSources, opts,
                        Sub(src)
                            Try
                                CleanupWithReferencesForModule(src.FCI)
                            Catch ex As Exception
                                LogError($"CleanupWithReferences({src.Name})", ex)
                            End Try
                        End Sub)
                Else
                    For Each src In GlobalSources
                        Try
                            CleanupWithReferencesForModule(src.FCI)
                        Catch ex As Exception
                            LogError($"CleanupWithReferences({src.Name})", ex)
                        End Try
                    Next
                End If
            End Sub

            runner()

        Catch ex As Exception
            LogError("CleanupWithReferences", ex)
        End Try
    End Sub

    ''' <summary>
    ''' Curăță referințele circulare pentru un modul individual (Form/Class/Module).
    ''' </summary>
    Private Sub CleanupWithReferencesForModule(modObj As ModuleContainer)
        If modObj Is Nothing Then Exit Sub

        ' === Curățare metode ===
        For Each m In modObj.Methods.Values
            For Each line In m.MethodLines
                ' 1️⃣ Elimină legături între linii (With, blocuri)
                line.WithBlockContext = Nothing
                line.EnumBlockContext = Nothing
                line.TypeBlockContext = Nothing
                ' 2️⃣ Elimină referințe din tokeni
                For Each t In line.Tokens
                    t.SourceLine = Nothing
                    t.WithBlockContext = Nothing
                    t.Parent = Nothing
                    't.ResolvedRef = Nothing
                Next
            Next
        Next

        ' === Curățare linii de la nivel de modul ===
        For Each line In modObj.Lines
            line.WithBlockContext = Nothing
            For Each t In line.Tokens
                t.SourceLine = Nothing
                t.WithBlockContext = Nothing
                t.Parent = Nothing
                't.ResolvedRef = Nothing
            Next
        Next
    End Sub

End Module
