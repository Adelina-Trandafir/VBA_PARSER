Namespace VBA_CORE
    Public Module modGlobals
        ' === CONFIGURAȚIE GLOBALĂ (înlocuiește argumentele CLI) ===
        Public Property Global_DBPath As String = "C:\AVACONT\Avacont.accdb"
        Public Property Global_UseCache As Boolean = False
        Public Property Global_UseParallel As Boolean = True
        Public Property Global_VerboseLevel As Integer = 1  ' 0=quiet,1=normal,2=verbose
        Public Property Global_ExportDir As String = "C:\ExportVBA"
    End Module
End Namespace
