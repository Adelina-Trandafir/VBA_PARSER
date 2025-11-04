'PROJECT NAME: VBA_CORE
'FILE DESCRIPTION: Modul pentru variabile și setări globale
'PATH: VBA_CORE/modGlobals.vb
Imports System.Reflection.Emit
Imports System.Runtime.Remoting.Messaging
Imports System.Threading
Imports VBA_CORE.modTypes

Public Module modGlobals
    ' ============================================================
    ' ENUM: TokenType
    ' Descriere: toate tipurile logice de tokeni din VBA
    ' ============================================================
    Public Enum TokenTypeEnum
        ' -----------------------------
        ' DECLARAȚII (introduc simboluri)
        ' -----------------------------
        module_decl                  ' modul, clasă, formă etc.
        variable_decl                ' Dim / Private / Public
        constant_decl                ' Const
        parameter_decl               ' parametru funcție/sub
        param_decl_api               ' parametru în Declare API
        function_decl                ' Function ...
        sub_decl                     ' Sub ...
        property_get_decl            ' Property Get ...
        property_let_decl            ' Property Let ...
        property_set_decl            ' Property Set ...
        event_decl                   ' Event ...
        enum_decl                    ' membru Enum
        type_decl                    ' Type ...
        type_field_decl              ' câmp în Type
        declare_function             ' Declare Function ...
        declare_sub                  ' Declare Sub ...

        ' -----------------------------
        ' SIMBOLURI REZOLVATE
        ' -----------------------------
        variable                     ' variabilă rezolvată
        constant                     ' constantă rezolvată
        parameter                    ' parametru rezolvat
        [Function]                     ' funcție internă
        subroutine                   ' sub intern
        property_get                 ' proprietate Get
        property_let                 ' proprietate Let
        property_set                 ' proprietate Set
        event_symbol                 ' eveniment rezolvat
        enum_type                    ' tip Enum
        enum_member                  ' membru Enum
        user_type                    ' tip definit (Type)
        type_field                   ' câmp în Type
        form_control                 ' control Access / UserForm
        form_property                ' proprietate control
        module_self                  ' instanța Me / modul curent
        module_symbol                ' simbol de nivel modul

        ' -----------------------------
        ' REFERINȚE ȘI EXPRESII
        ' -----------------------------
        type_ref                     ' referință la un tip
        object_ref                   ' referință la obiect
        variable_member              ' membru al unei variabile (obj.Prop)
        function_call                ' apel funcție
        property_access              ' acces proprietate
        method_call                  ' apel metodă
        with_member                  ' membru din bloc With
        member_chain                 ' parte din lanț a.b.c
        index_access                 ' acces cu paranteze (arr(0))
        unknown_type                 ' tip necunoscut după As
        unresolved                   ' token nerecunoscut
        self                         ' token "Me"

        ' -----------------------------
        ' BUILT-IN & ACCESS SPECIFIC
        ' -----------------------------
        builtin_type                 ' tip VBA (String, Long)
        builtin_function             ' funcție VBA (Trim, Left)
        builtin_constant             ' constantă VBA (vbYes)
        builtin_object               ' obiect VBA (Application, DoCmd)
        access_function              ' funcție Access (Nz, DLookup)
        access_object                ' obiect Access (Forms, Reports)
        builtin_property             ' proprietate built-in (.Name, .Value)
        vba_keyword                  ' cuvânt cheie (If, End, For)

        ' -----------------------------
        ' CONTEXT / STRUCTURĂ / CONTROL
        ' -----------------------------
        event_handler                ' procedură de eveniment
        label                        ' etichetă (MyLabel:)
        line_continuation            ' linie cu _
        comment                      ' comentariu '
        preprocessor                 ' #If, #Const etc.
        block_start                  ' început bloc (If, For)
        block_end                    ' sfârșit bloc (End If, Next)
        with_context                 ' context activ With
        error_handler                ' On Error / Resume
        call_statement               ' apel explicit Call

        ' -----------------------------
        ' ERORI / DIAGNOSTICE
        ' -----------------------------
        syntax_error                 ' linie invalidă
        invalid_token                ' token invalid
        ambiguous_ref                ' referință ambiguă
        shadowed_symbol              ' simbol mascat
        duplicate_decl               ' dublă declarație
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
End Module
