'PROJECT NAME: VBA_CORE
'FILE DESCRIPTION: Modul pentru variabile și setări globale
'PATH: VBA_CORE/modGlobals.vb
Imports System.Reflection.Emit
Imports System.Runtime.Remoting.Messaging
Imports System.Threading
Imports VBA_CORE.customTypes

Public Module modGlobals
    ''' <summary>
    ''' Tipuri logice de tokeni (enumerație unificată pentru analizorul VBA).
    ''' </summary>
    Public Enum TokenTypeEnum
        ' === 0️⃣ Generic ===
        unresolved
        type_ref
        self
        initial

        ' === 1️⃣ Declarații ===
        variable_decl
        constant_decl
        parameter_decl
        function_decl
        sub_decl
        property_get_decl
        property_let_decl
        property_set_decl
        event_decl
        declare_function
        declare_sub
        type_member_decl
        enum_member_decl
        label_decl
        label_ref

        ' === 2️⃣ Utilizări / referințe ===
        variable
        constant
        parameter
        [Function]
        [sub]
        property_get
        property_let
        property_set
        event_handler
        variable_member
        builtin_property

        ' === 3️⃣ Tipuri complexe ===
        enum_type
        enum_member
        user_type
        type_field

        ' === 4️⃣ Built-in și Access ===
        builtin_type
        builtin_function
        builtin_object
        builtin_constant
        access_object
        access_function
        access_property
        access_event

        ' === 5️⃣ Altele ===
        vba_keyword
        external_prefix

        ' === 6️⃣ Access Form/Report Objects ===
        access_control
        access_section
        access_form
        access_report
    End Enum

    ' ===============================
    ' 🧰 CONFIG RUNTIME
    ' ===============================
    Public mainThreadId As Integer = Thread.CurrentThread.ManagedThreadId


    ' === CONFIGURAȚIE GLOBALĂ (înlocuiește argumentele CLI) ===
    Public Property Global_DBPath As String = "" '"C:\AVACONT\Avacont.accdb"
    Public Property Global_UseCache As Boolean = False
    Public Property Global_UseParallel As Boolean = False
    Public Property Global_VerboseLevel As Integer = 1  ' 0=quiet,1=normal,2=verbose
    Public Property Global_ExportDir As String = "C:\ExportVBA"
    Public Property ExtendedVerbose As Boolean = False
End Module
