
'PROJECT NAME: VBA_CORE
'FILE DESCRIPTION: Modul pentru regex-uri utilizate în analiza VBA
'PATH: VBA_CORE/modRegexHelper.vb
Imports System.Text.RegularExpressions

Public Class RegexCache
    ' Main parsing
    ' Method Tester
    Public Shared ReadOnly RxMethodTester As New Regex("^\s*(Public|Private|Friend)?\s*(Function|Sub|Property)\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    ' Method definitions
    Public Shared ReadOnly RxHeader As New Regex("^(?:(Public|Private|Friend)\s+)?(?:(Function|Sub))\s+([A-Za-z_]\w*[$%!#&@]?)\s*\(\s*((?:[^()""]+|""(?:[^""]|"""")*""|\([^()]*\))*)\s*\)(?:\s+As\s+([A-Za-z_]\w*(?:\.[A-Za-z_]\w*)*(?:\(\))?))?", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    ' Events
    Public Shared ReadOnly RxEvent As New Regex("^(?:(Public|Private|Friend)\s+)?Event\s+([A-Za-z_]\w*)\s*\(\s*((?:[^()""]+|""(?:[^""]|"")*""|\([^()]*\))*)\s*\)", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    'Properties
    Public Shared ReadOnly RxProperty As New Regex("^(?:(Public|Private|Friend)\s+)?Property\s+(Get|Let|Set)\s+([A-Za-z_]\w*)\s*\(\s*((?:[^()""]+|""(?:[^""]|"")*""|\([^()]*\))*)\s*\)(?:\s+As\s+([A-Za-z_]\w*(?:\.[A-Za-z_]\w*)*(?:\(\))?))?", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    ' Tipul sectiunii din Form/Report
    Public Shared ReadOnly RxSectionType As New Regex("^(Section|FormHeader|FormFooter|ReportHeader|ReportFooter|PageHeader|PageFooter|FormPageHeader|FormPageFooter|ReportPageHeader|ReportPageFooter)$", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    ' Decteaza inceputul unei machete
    Public Shared ReadOnly RxFormOrReport As New Regex("^(Form|Report)$", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    'Decteaza inceputul unei sectiuni de formular/raport
    Public Shared ReadOnly RxBeginPrtForm As New Regex("^(?i)\s*([A-Za-z_]\w*)\s*=\s*Begin\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    Public Shared ReadOnly RxEndMethod As New Regex("^\s*End\s+(Sub|Function|Property)\s*$", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public Shared ReadOnly RxParamsInit As New Regex("\((.*)\)", RegexOptions.Singleline Or RegexOptions.Compiled)
    'detectie parametrii in semnatura metodei
    Public Shared ReadOnly RxParams As New Regex("(?:(ByVal|ByRef)\s+)?(?:(Optional)\s+)?(?:(ParamArray)\s+)?([A-Za-z_]\w*(?:\s*\(\))?)\s*(?:As\s+([A-Za-z_]\w*(?:\.[A-Za-z_]\w*)*(?:\(\))?))?(?:\s*=\s*("".*?""|[^,]+))?\s*(?:,|$)", RegexOptions.Compiled)

    Public Shared ReadOnly RxAsType As New Regex("\bAs\s+([A-Za-z_][A-Za-z0-9_\.]*)", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public Shared ReadOnly RxNew As New Regex("(?i)\bSet\s+([A-Za-z_]\w*)\s*=\s*New\s+([A-Za-z_][A-Za-z0-9_\.]*)", RegexOptions.Compiled)

    ' 🔥 ÎMBUNĂTĂȚIT: Regex-uri pentru referințe (vor fi aplicate pe linii fără string-uri)
    Public Shared ReadOnly RxCallObj As New Regex("(?<![A-Za-z0-9_])([A-Za-z_]\w*)\s*\.\s*([A-Za-z_]\w*)\s*\(", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public Shared ReadOnly RxPropObj As New Regex("(?<![A-Za-z0-9_])([A-Za-z_]\w*)\s*\.\s*([A-Za-z_]\w*)\b(?!\s*\()", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public Shared ReadOnly RxCallBare As New Regex("(?<![A-Za-z0-9_])([A-Za-z_]\w*)\s*\(", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    ' Declares
    Public Shared ReadOnly RxDeclares As New Regex("^(?i)(Public|Private)?\s*(Declare(?:\s+PtrSafe)?)\s+(Function|Sub)\s+([A-Za-z_]\w*)\s+Lib\s+""([^""]+)""(?:\s+Alias\s+""([^""]+)"")?\s*\(\s*((?:[^()""']+|""(?:[^""]|"")*""|\([^()]*\))*)\s*\)\s*(?:As\s+([\w.]+))?", RegexOptions.Compiled)

    ' Tokenizer
    'Public Shared ReadOnly RxTokenizer As New Regex("([A-Za-z_]\w*)\s*(=|\(|\.|,|\+|\-|\*|\/|&|\)|\s+(?=[A-Za-z_]\w*)|$)(?=(?:[^""]*""[^""]*"")*[^""]*$)", RegexOptions.Compiled)
    Public Shared ReadOnly RxTokenizer As New Regex("(\.?[A-Za-z_]\w*(?:\.[A-Za-z_]\w*)*)(?!(?:[^<]*>))\s*(=|\(|,|\+|\-|\*|/|&|\)|As|\s+|$)", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    ' Split dupa :
    Public Shared ReadOnly RxSplitColon As New Regex(":(?=(?:[^""]*""[^""]*"")*[^""]*$)", RegexOptions.Compiled)

    ' Variable detection helpers
    Public Shared ReadOnly RxLocalVar As New Regex("(?i)^\s*(?:Dim|Static)\s+([A-Za-z_]\w*)\s*(?:As\s+(?:New\s+)?([A-Za-z_][\w\.]*))?", RegexOptions.Compiled)

    Public Shared ReadOnly RxClassField As New Regex("(?i)^\s*(?:Public|Private)\s+(?<we>WithEvents\s+)?(?<n>[A-Za-z_]\w*)\s*(?:As\s+(?<type>[A-Za-z_][\w\.]*))?", RegexOptions.Compiled)

    'detectie parametrii in modul/clasa
    Public Shared ReadOnly RxGlobalVar As New Regex("(?i)^\s*(Private|Public|Global|Dim)\s+(?<we>WithEvents\s+)?(?<n>[A-Za-z_]\w*)\s*(?:As\s+(?<type>[A-Za-z_][\w\.]*))?", RegexOptions.Compiled)

    Public Shared ReadOnly RxPrimitiveType As New Regex("^(Integer|Long|String|Boolean|Double|Variant|Object|Byte|Currency|Decimal|Single)?$", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    ' 🔥 Tokenizare - extrage identificatori (va fi aplicat pe linii curate)
    Public Shared ReadOnly RxIdentifier As New Regex("[A-Za-z_][A-Za-z0-9_]*", RegexOptions.Compiled)

    ' 🔥 Detectare literali pentru curățare
    Public Shared ReadOnly RxDateLiteral As New Regex("#[^#]+#", RegexOptions.Compiled)

    ' 🔥 Detectare proprietăți formulare/controale
    Public Shared ReadOnly RxName As New Regex("(?im)^\s*Name\s*=\s*(.+)$", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public Shared ReadOnly RxCtrlSrc As New Regex("(?im)^\s*ControlSource\s*=\s*(.+)$", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public Shared ReadOnly RxTag As New Regex("(?im)^\s*Tag\s*=\s*(.+)$", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public Shared ReadOnly RxRecSrc As New Regex("(?im)^\s*RecordSource\s*=\s*(.+)$", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    ' Detectie comentarii linie (doar daca nu sunt intre "")
    Public Shared ReadOnly RxComments As New Regex("^([^'""]*(?:""[^""]*""[^'""]*)*)'", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public Shared ReadOnly RxSpaces As New Regex("\s{2,}", RegexOptions.Compiled)

    ' Detectie constante
    Public Shared ReadOnly RxConsts As New Regex("^(?:(Public|Private|Friend)\s+)?(Const)\s+([A-Za-z_]\w*)(?:\s+As\s+([\w.]+))?\s*=\s*(.+)$", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    ' Keywords VBA (pentru filtrare) - 🔥 ÎMBUNĂTĂȚIT: adăugate mai multe
    Public Shared ReadOnly VbaKeywords As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
        "Sub", "Function", "Property", "End", "Private", "Public", "Friend", "Static", "Dim",
        "As", "Integer", "Long", "String", "Boolean", "Double", "Variant", "Object", "Date",
        "Byte", "Currency", "Decimal", "Single",
        "If", "Then", "Else", "ElseIf", "End If", "Select", "Case", "For", "Next", "Do", "Loop",
        "While", "Wend", "Exit", "Return", "Call", "Set", "Let", "New", "Nothing", "Null",
        "True", "False", "And", "Or", "Not", "Xor", "Mod", "Like", "Is", "To", "Step",
        "ByVal", "ByRef", "Optional", "ParamArray", "Const", "Enum", "Type", "With", "ReDim",
        "Preserve", "Erase", "On", "Error", "Resume", "GoTo", "GoSub", "Me", "MyBase", "MyClass",
        "Each", "In", "Until", "Get", "Put", "Open", "Close", "Input", "Output", "Binary",
        "Random", "Append", "Lock", "Unlock", "Seek", "Print", "Write", "Line", "Width",
        "Implements", "Handles", "Inherits", "MustInherit", "MustOverride", "Overrides",
        "Overridable", "NotOverridable", "Overloads", "Shadows", "Shared", "Protected"
    }

    ' Detectie Type/Enum
    Public Shared ReadOnly RxEnumTypeStart As New Regex("(?:Type |Enum )(\w+)", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public Shared ReadOnly RxEnumMember As New Regex("^(?:\[?(?<name>[A-Za-z_]\w*|\[[^\]]+\])\]?)(?:\s*=\s*(?<value>[-+]?(?:&H[0-9A-Fa-f]+|\d+)))?", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public Shared ReadOnly RxTypeMember As New Regex("^(?<name>[A-Za-z_]\w*)\s+As\s+(?<type>[A-Za-z_]\w*(?:\.[A-Za-z_]\w*)?(?:\([^)]*\))?)", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    ' Detectie LABEL-uri de linie (GoTo s)
    Public Shared ReadOnly RxCodeLabel As New Regex("^[A-Za-z_]\w*:$", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    ' Designer
    Public Shared ReadOnly RxDesigner As New Regex("(""(?:[^""\\]|\\.)*""|(!=)|[^\s""]+)", RegexOptions.IgnoreCase Or RegexOptions.Compiled Or RegexOptions.Singleline)

    ' Line cleaner (removing strings and comments)
    Public Shared ReadOnly RxLineCleaner As New Regex("(""([^""\\]|\\.)*"")|('.*$)", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    ' Pretokenizer
    Public Shared ReadOnly RxPreTokenizer As New Regex("(?>\<[^>]*\>)|(\s*[A-Za-z_]\w*(?:[.!][A-Za-z_]\w*)*(?:\s+As\s+[A-Za-z_]\w*)?)", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public Shared ReadOnly RxAsBlock As New Regex("(?<asblock>\b(?<name>[A-Za-z_]\w*)\s+As\s+(?<datatype>[A-Za-z_]\w*(?:\.[A-Za-z_]\w*)*))", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
End Class
