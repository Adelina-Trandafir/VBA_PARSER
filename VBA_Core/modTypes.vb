Imports System.Collections.Concurrent

' ==========================================================
'  MODEL DE DATE GLOBAL – vizibil din toate modulele
' ==========================================================
Public Module modTypes
    <Serializable>
    Public Class VariableInfo
        Public Name As String
        Public TypeName As String
        Public CallerModule As String
        Public ScopeFunc As String              ' "" => la nivel modul
        Public IsWithEvents As Boolean
    End Class

    <Serializable>
    Public Class Token
        Public Property Parent As Token
        Public Property LineNumber As Integer
        Public Property TokenString As String
        Public Property TokenType As String      ' "identifier", "keyword", "operator", "literal"
        Public Property Context As String        ' numele metodei (ex: "Command37_Click")
        Public Property NextSymbol As String
        Public Property IsLeftSide As Boolean
        Public Property DataType As String      ' tipul variabilei
        Public Property ParentType As String    ' tipul obiectului părinte
    End Class

    <Serializable>
    Public Class ModuleContainer
        Public Property Name As String
        Public Property Type As String      ' "Form" / "Class" / "Module"
        Public Property CodeContent As String
        Public Property Methods As New Dictionary(Of String, MethodInfo)
        Public Property Lines As New List(Of MethodLine)
        Public Property FilePath As String
        Public Property Constants As New List(Of ParamInfo)
        Public Property Variables As New List(Of ParamInfo)
        Public Property Declares As New Dictionary(Of String, MethodInfo)
    End Class

    <Serializable>
    Public Class FormReportContainer
        Public Property Name As String
        Public Property Type As String      ' "Form" / "Report"
        Public Property CodeContent As String
        Public Property Methods As New Dictionary(Of String, MethodInfo)
        Public Property Lines As New List(Of MethodLine)
        Public Property FilePath As String
        Public Property RecordSource As String
        Public Property Tag As String
        Public Property Sections As New List(Of FormSection)
        Public Property Controls As New List(Of FormControl)
        Public Property Constants As New List(Of ParamInfo)
        Public Property Variables As New List(Of ParamInfo)
        Public Property Declares As New Dictionary(Of String, MethodInfo)
    End Class

    <Serializable>
    Public Class FormSection
        Public Property Name As String
        Public Property Type As String    ' ex: "Section", "FormHeader", "ReportFooter"
        Public Property Controls As New List(Of FormControl)
        Public Property Parent As New FormReportContainer
    End Class

    <Serializable>
    Public Class FormControl
        Public Property Type As String = ""
        Public Property Name As String = ""
        Public Property ControlSource As String = ""
        Public Property RecordSource As String = ""
        Public Property RowSource As String = ""
        Public Property SourceObject As String = ""
        Public Property Caption As String = ""
        Public Property ParentSection As String = ""
        Public Property Section As FormSection
        Public Property Tag As String = ""
        Public Property SubLabel As FormControl

        Public Overrides Function ToString() As String
            Dim props As New List(Of String)
            If ControlSource <> "" Then props.Add($"ControlSource={ControlSource}")
            If RecordSource <> "" Then props.Add($"RecordSource={RecordSource}")
            If RowSource <> "" Then props.Add($"RowSource={RowSource}")
            If SourceObject <> "" Then props.Add($"SourceObject={SourceObject}")
            If Caption <> "" Then props.Add($"Caption={Caption}")
            Dim p = If(props.Count > 0, " → " & String.Join(", ", props), "")
            Dim lbl = If(SubLabel IsNot Nothing, $" [Lbl:{SubLabel.Name}]", "")
            Return $"{Type}: {Name}{p}{lbl}"
        End Function
    End Class

    <Serializable>
    Public Class MethodInfo
        Public Property Name As String = ""          ' ex: Command37_Click
        Public Property Signature As String = ""     ' linia header-ului
        Public Property ParentModule As String = ""  ' nume modul (ex: frmMain)
        Public Property ParentType As String = ""    ' Modules / Forms / Reports / Class
        Public Property StartLine As Integer = 0
        Public Property EndLine As Integer = 0
        Public Property MethodScope As String = ""   ' Private, Public, Friend
        Public Property MethodLines As New List(Of MethodLine)
        Public Property Tokens As New List(Of Token)  ' toți identificatorii din metodă
        Public Property IsFunction As Boolean
        Public Property IsDefault As Boolean        ' Attribute Item.VB_UserMemId = 0
        Public Property IsProperty As Boolean
        Public Property IsEnumerable As Boolean     ' Attribute NewEnum.VB_MemberFlags = "40"
        Public Property ReturnType As String
        Public Property ReturnObject As VariableInfo
        Public Property Parameters As New List(Of ParamInfo)
        Public Property Constants As New List(Of ParamInfo)
        Public Property Variables As New List(Of ParamInfo)
        Public Property Events As New List(Of MethodLine)
        Public Property TokenizedLine As MethodLine
        Public Property CodeContent As String
        Public Property PropertyType As String      ' pentru proprietăți: Get / Let / Set
    End Class

    <Serializable>
    Public Class MethodLine
        Public Property LineNumber As Integer
        Public Property Content As String
        Public Property Tokens As New List(Of Token)
    End Class

    <Serializable>
    Public Class LocalSymbolInfo
        Public Property Name As String              ' numele variabilei/funcției
        Public Property DeclaringModule As String   ' modulul unde e declarat
        Public Property DeclaringScope As String    ' funcția/clasa unde e declarat
        Public Property SymbolType As String        ' "variable", "function", "property", "sub"
        Public Property DataType As String          ' tipul de date (pentru variabile)
        Public Property IsPrivate As Boolean
    End Class

    <Serializable>
    Public Class NonStandardCallInfo
        Public Property CallerModule As String
        Public Property CallerMethod As String
        Public Property CalleeToken As String       ' token-ul apelat (ex: "obj.Method" sau "Method")
        Public Property LineNumber As Integer
        Public Property IsQualified As Boolean      ' True dacă e "obj.Method", False dacă e "Method"
    End Class

    <Serializable>
    Public Class ParamInfo
        Public Property Name As String
        Public Property Type As String
        Public Property IsGlobal As Boolean
        Public Property IsPublic As Boolean
        Public Property IsByRef As Boolean
        Public Property IsOptional As Boolean
        Public Property IsParamArray As Boolean
        Public Property IsWithEvents As Boolean
        Public Property ParentMethod As Object
        Public Property Value As String
        Public Property Tokens As List(Of Token)
    End Class

    ' Liste globale pentru clasificare token-uri
    'Public GlobalLocalSymbols As New Concurrent.ConcurrentBag(Of LocalSymbolInfo)
    'Public GlobalNonStandardCalls As New Concurrent.ConcurrentBag(Of NonStandardCallInfo)
    Public GlobalModules As New Dictionary(Of String, ModuleContainer)(StringComparer.OrdinalIgnoreCase)
    Public GlobalFormsReports As New Dictionary(Of String, FormReportContainer)(StringComparer.OrdinalIgnoreCase)

    ' === Clasă helper pentru context (thread-safe collections) ===
    'Public Class ClassificationContext
    '    Public Property GlobalPublicSymbols As HashSet(Of String)
    '    Public Property FormControlsIndex As Concurrent.ConcurrentDictionary(Of String, HashSet(Of String))
    '    Public Property LocalVarsIndex As Concurrent.ConcurrentDictionary(Of String, Dictionary(Of String, HashSet(Of String)))
    '    Public Property PrivateMethodsIndex As Concurrent.ConcurrentDictionary(Of String, HashSet(Of String))
    'End Class

    'Public GlobalPrivateMethods As New Concurrent.ConcurrentDictionary(Of String, List(Of MethodInfo))(StringComparer.OrdinalIgnoreCase)

    'Public GlobalSymbols As New ConcurrentQueue(Of SymbolInfo)
    'Public GlobalReferences As New ConcurrentQueue(Of ReferenceInfo)
    'Public GlobalVariables As New ConcurrentQueue(Of VariableInfo)


    'Public Class FunctionInfo
    '    Public Property Signature As String = ""
    '    Public Property Name As String = ""
    '    Public Property ParentClass As String = ""
    '    Public Property StartLine As Integer = 0
    '    Public Property EndLine As Integer = 0
    '    Public Property LineCount As Integer = 0
    '    Public Property Complexity As Integer = 0
    '    Public Property CallSites As New List(Of String)
    '    Public Property CalledBy As New List(Of String)
    '    Public Property Accessibility As String = "Public"
    'End Class

    'Public Class ExportedModule
    '    Public Property Name As String
    '    Public Property Type As String
    '    Public Property FilePath As String
    '    Public Property Lines As String()
    '    Public Property CleanLines As String()

    '    Public Property Functions As New List(Of FunctionInfo)
    '    Public Property References As New List(Of String)
    '    Public Property OutgoingModules As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    '    Public Property IncomingModules As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    'End Class

    'Public Class PreloadedFile
    '    Public Property Path As String
    '    Public Property Content As String
    '    Public Property Lines As String()
    '    Public Property CleanLines As String()
    '    Public Property Type As String    ' Module / Form / Report / Macro

    'End Class

    'Public Class SymbolInfo
    '    Public Name As String                    ' ex: SaveClient (nume simplu)
    '    Public DeclaringModule As String         ' ex: modClients, frmMain
    '    Public DeclaringClass As String          ' ex: clsClient ("" dacă nu e în clasă)
    '    Public IsPublic As Boolean
    '    Public IsProperty As Boolean
    '    Public IsFunction As Boolean
    '    Public IsSub As Boolean
    '    Public ReturnType As String              ' ex: "Long", "" pentru Sub / nedefinit
    '    Public Parameters As New List(Of String) ' parametri raw (text)
    '    Public Parent As Object
    'End Class

    'Public Class ReferenceInfo
    '    Public CallerModule As String
    '    Public CallerClass As String
    '    Public CallerFunc As String             ' funcția curentă (completă: Class.Method sau Method)
    '    Public LineNumber As Integer
    '    Public Token As String                  ' ex: "a.Save", "Save", "a.Prop"
    'End Class

    ' ==========================================================
    '  TOKEN CLASSIFICATION & DEPENDENCY ANALYSIS
    ' ==========================================================
    'Public Class TokenClassification
    '    Public Property Token As String
    '    Public Property Category As String          ' "keyword", "control", "local_private", "local_public", "global_public"
    '    Public Property DeclaringScope As String    ' numele funcției/clasei unde e declarat
    '    Public Property DeclaringModule As String   ' numele modulului
    '    Public Property LineNumber As Integer
    '    Public Property Context As String           ' metoda unde apare token-ul
    'End Class
End Module