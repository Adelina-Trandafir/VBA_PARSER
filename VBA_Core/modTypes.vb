'PROJECT NAME: VBA_CORE
'FILE DESCRIPTION: Modul pentru tipuri de date utilizate în analiza VBA
'PATH: VBA_CORE/modTypes.vb

' ==========================================================
'  MODEL DE DATE GLOBAL – vizibil din toate modulele
' ==========================================================
Imports System.Collections.Concurrent

Public Module modTypes
    <Serializable>
    Public Class ModuleContainer
        Public Property Name As String
        Public Property Type As String      ' "Form" / "Class" / "Module"
        Public Property CodeContent As String
        Public Property FilePath As String
        Public Property Lines As New List(Of MethodLine)
        Public Property Constants As New List(Of ParamInfo)
        Public Property Variables As New List(Of ParamInfo)
        Public Property Types As New List(Of EnumOrType)
        Public Property Enums As New List(Of EnumOrType)
        Public Property Declares As New Dictionary(Of String, MethodInfo)
        Public Property Events As New Dictionary(Of String, MethodInfo)
        Public Property Methods As New Dictionary(Of String, MethodInfo)
        Public Property Properties As New Dictionary(Of String, PropertyInfo)
        Public Property Functions As New Dictionary(Of String, FunctionInfo)
        Public Property ModuleSymbols As New List(Of SymbolEntry)

        ' =============================================================
        ' 🔹 Helper pentru identificarea rapidă a simbolurilor
        ' =============================================================
        Public Function GetMemberInfo(name As String) As (Kind As String, ReturnType As String)
            If String.IsNullOrEmpty(name) Then Return ("", "")

            ' === Funcție ===
            Dim f As FunctionInfo = Nothing
            If Functions.TryGetValue(name, f) Then
                Return ("Function", f.ReturnType)
            End If

            ' === Proprietate ===
            Dim p As PropertyInfo = Nothing
            If Properties.TryGetValue(name, p) Then
                Return ($"Property_{p.PropertyType}", p.ReturnType) ' ex: Property_Get / Property_Let / Property_Set
            End If

            ' === Variabilă ===
            Dim v = Variables.FirstOrDefault(Function(x) x.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
            If v IsNot Nothing Then
                Return ("Variable", v.Type)
            End If

            ' === Constantă ===
            Dim c = Constants.FirstOrDefault(Function(x) x.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
            If c IsNot Nothing Then
                Return ("Constant", c.Type)
            End If

            Return ("", "")
        End Function
    End Class

    <Serializable>
    Public Class EnumOrType
        Public Property Name As String
        Public Property Type As String      ' "Enum" / "Type"
        Public Property IsPublic As Boolean
        Public Property IsGlobal As Boolean
        Public Property Members As New List(Of ParamInfo)
    End Class

    <Serializable>
    Public Class FormReportContainer
        Inherits ModuleContainer

        Public Property RecordSource As String
        Public Property Tag As String
        Public Property Sections As New List(Of FormSection)
        Public Property Controls As New List(Of FormControl)
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
        Public Property Name As String = ""
        Public Property Signature As String = ""
        Public Property ParentModule As String = ""
        Public Property ParentType As String = ""
        Public Property StartLine As Integer = 0
        Public Property EndLine As Integer = 0
        Public Property MethodScope As String = ""   ' Public, Private, Friend
        Public Property MethodLines As New List(Of MethodLine)
        Public Property Tokens As New List(Of Token)
        Public Property CodeContent As String
        Public Property IsDefault As Boolean
        Public Property IsEnumerable As Boolean
        Public Property Parameters As New List(Of ParamInfo)
        Public Property Constants As New List(Of ParamInfo)
        Public Property Variables As New List(Of ParamInfo)
        Public Property TokenizedLine As MethodLine
        Public Property IsAccessEventHandler As Boolean
        Public Property HandlerObject As String
        Public Property HandlerEvent As String
        Public Property HandlerObjectType As String

    End Class

    Public Class FunctionInfo
        Inherits MethodInfo
        Public Property IsFunction As Boolean
        Public Property ReturnType As String
        Public Property ReturnObject As ParamInfo
    End Class

    Public Class PropertyInfo
        Inherits MethodInfo
        Public Property IsProperty As Boolean
        Public Property ReturnType As String
        Public Property ReturnObject As ParamInfo
        Public Property PropertyType As String      ' pentru proprietăți: Get / Let / Set
    End Class

    <Serializable>
    Public Class MethodLine
        Public Property LineNumber As Integer
        Public Property LocalLineNumber As Integer
        Public Property Content As String
        Public Property OriginalContent As String
        Public Property Tokens As New List(Of Token)
        '------ ENUM METHOD BLOCK CONTEXT ------
        Public Property StartsEnumBlock As Boolean
        Public Property IsInEnumBlock As Boolean
        Public Property EndsEnumBlock As Boolean
        Public Property EnumBlockContext As MethodLine
        '------ TYPE METHOD BLOCK CONTEXT ------
        Public Property StartsTypeBlock As Boolean
        Public Property IsInTypeBlock As Boolean
        Public Property EndsTypeBlock As Boolean
        Public Property TypeBlockContext As MethodLine
        '------ FUNCTION / SUB / PROPERTY METHOD BLOCK CONTEXT ------
        Public Property StartsMethodBlock As Boolean
        Public Property IsDeclarationLine As Boolean
        Public Property EndsMethodBlock As Boolean
        '------ COND IF BLOCK CONTEXT ------
        Public Property StartsCondIfBlock As Boolean
        Public Property EndsCondIfBlock As Boolean
        '------ LINE MODIFICATION FLAGS ------
        Public Property WasSplit As Boolean
        Public Property WasConcat As Boolean
        '------ WITH BLOCK CONTEXT ------
        Public Property StartsWithBlock As Boolean
        Public Property IsInWithBlock As Boolean
        Public Property EndsWithBlock As Boolean
        Public Property WithBlockContext As MethodLine
        Public Property WithBlockDepth As Integer
        '------ EVENTS ------
        Public Property IsEventLine As Boolean
    End Class

    <Serializable>
    Public Class Token
        Public Property Parent As Token
        Public Property LineNumber As Integer
        Public Property ColumnNumber As Integer
        Public Property LocalLineNumber As Integer ' linia relativă în metoda curentă
        Public Property LocalColumnNumber As Integer ' 
        Public Property MethodLineNumber As Integer ' linia relativă în metoda curentă
        Public Property MethodColumnNumber As Integer ' 
        Public Property TokenString As String
        Public Property TokenType As TokenTypeEnum
        Public Property Context As String        ' numele metodei (ex: "Command37_Click")
        Public Property NextSymbol As String
        Public Property IsLeftSide As Boolean
        Public Property IsReturnValue As Boolean
        Public Property DataType As String      ' tipul variabilei
        Public Property ParentType As String    ' tipul obiectului părinte
        Public Property ResolvedRef As Object
        Public Property IsResolved As Boolean
        Public Property IsWithMember As Boolean ' flag dacă e membru al unui With
        Public Property WithBlockContext As MethodLine
        Public Property WithBlockDepth As Integer
        Public Property QualifiedName As String
        Public Property SourceLine As MethodLine
        Public Property IsBuiltIn As Boolean
        Public Property ResolvedScope As String ' "local", "global", "with_context", etc.

        Public Overrides Function ToString() As String
            Dim props As New List(Of String)
            If TokenString <> "" Then props.Add($"Token='{TokenString}'")
            If TokenType <> "" Then props.Add($"Type={TokenType}")
            If DataType <> "" Then props.Add($"DataType={DataType}")
            props.Add($"Global_Pos={LineNumber}:{ColumnNumber}")
            props.Add($"Method_Pos={MethodLineNumber}:{MethodColumnNumber}")
            Dim p = If(props.Count > 0, String.Join(";", props), "")
            Return p
        End Function
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
        Public Property DataType As String
        Public Property DeclaredInLine As Integer
        Public Property Scope As String      ' "Public", "Private", "Dim", etc.
        Public Property QualifiedName As String
        Public Property Tokens As List(Of Token)
    End Class

    Public Class SymbolEntry
        Public Property Name As String
        Public Property TypeName As String
        Public Property DeclType As String   ' "Variable", "Const", "Function", "Property"
        Public Property Scope As String      ' "Public", "Private", "Friend"
        Public Property SourceObject As Object ' referință către obiectul real (ParamInfo, FunctionInfo, PropertyInfo)
        Public Property Category As String  ' "Local", "Global", "FormControl", etc.    
    End Class

    ' Liste globale pentru clasificare token-uri
    Public GlobalModules As New ConcurrentDictionary(Of String, ModuleContainer)(StringComparer.OrdinalIgnoreCase)
    Public GlobalFormsReports As New ConcurrentDictionary(Of String, FormReportContainer)(StringComparer.OrdinalIgnoreCase)

    Public GlobalSources As New ConcurrentBag(Of (Name As String, Type As String, Code As String, FCI As ModuleContainer))

    Public GlobalSymbols As New ConcurrentDictionary(Of String, SymbolEntry)(StringComparer.OrdinalIgnoreCase)
    Public Function FindSymbol(container As ModuleContainer, symbolName As String) As SymbolEntry
        If container?.ModuleSymbols Is Nothing Then Return Nothing
        Return container.ModuleSymbols.FirstOrDefault(Function(s) s.Name.Equals(symbolName, StringComparison.OrdinalIgnoreCase))
    End Function
End Module