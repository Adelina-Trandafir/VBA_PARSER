'PROJECT NAME: VBA_CORE
'FILE DESCRIPTION: Module pentru liste de simboluri built-in care trebuie ignorate
'PATH: VBA_CORE/modIgnoreLists.vb

Imports System.Text.RegularExpressions

Public Module modIgnoreLists
    ' ============================================================
    '  1. VBA BUILT-IN TYPES
    ' ============================================================
    ' Tipuri primitive        
    ' Collection types        
    ' Special

    Public ReadOnly VbaBuiltInTypes As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
        "Boolean", "Byte", "Integer", "Long", "LongLong", "LongPtr",
        "Single", "Double", "Currency", "Decimal",
        "Date", "String",
        "Object", "Variant",
        "Collection",
        "Any", "Empty", "Null", "Nothing"
    }

    ' ============================================================
    '  2. VBA BUILT-IN FUNCTIONS & STATEMENTS
    ' ============================================================
    ' String functions
    ' Conversion functions
    ' Math functions        
    ' Date/Time functions        
    ' Array functions        
    ' Type checking        
    ' File I/O        
    ' Interaction        
    ' Information        
    ' Conditional        
    ' Misc        
    ' Financial        
    ' Error handling

    Public ReadOnly VbaBuiltInFunctions As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
        "Asc", "AscB", "AscW", "Chr", "ChrB", "ChrW",
        "InStr", "InStrB", "InStrRev",
        "LCase", "UCase", "LCase$", "UCase$",
        "Left", "Right", "Mid", "Left$", "Right$", "Mid$",
        "Len", "LenB",
        "LTrim", "RTrim", "Trim", "LTrim$", "RTrim$", "Trim$",
        "Space", "Space$", "String", "String$",
        "Replace", "Split", "Join",
        "StrComp", "StrConv", "StrReverse",
        "Format", "Format$",
        "CBool", "CByte", "CCur", "CDate", "CDbl", "CDec", "CInt", "CLng", "CLngLng", "CLngPtr",
        "CSng", "CStr", "CVar",
        "Val", "Str", "Str$",
        "Hex", "Hex$", "Oct", "Oct$",
        "Abs", "Atn", "Cos", "Exp", "Fix", "Int",
        "Log", "Rnd", "Sgn", "Sin", "Sqr", "Tan",
        "Date", "Date$", "Time", "Time$",
        "Now", "Timer",
        "DateAdd", "DateDiff", "DatePart", "DateSerial", "DateValue",
        "TimeSerial", "TimeValue",
        "Day", "Month", "Year", "Weekday", "Hour", "Minute", "Second",
        "MonthName", "WeekdayName",
        "Array", "Filter", "IsArray", "LBound", "UBound",
        "IsDate", "IsEmpty", "IsError", "IsMissing", "IsNull", "IsNumeric", "IsObject",
        "TypeName", "VarType",
        "Dir", "Dir$", "EOF", "FileAttr", "FileDateTime", "FileLen",
        "FreeFile", "Input", "Input$", "InputB", "InputB$",
        "Loc", "LOF", "Seek",
        "MsgBox", "InputBox", "InputBox$",
        "Beep", "Shell",
        "CreateObject", "GetObject",
        "Command", "Command$", "Environ", "Environ$",
        "RGB", "QBColor",
        "IIf", "Choose", "Switch",
        "DDB", "FV", "IPmt", "IRR", "MIRR", "NPer", "NPV", "Pmt", "PPmt", "PV", "Rate", "SLN", "SYD",
        "Error", "Error$", "CVErr",
        "DoEvents", "SendKeys", "AppActivate",
        "CallByName", "Partition"
    }

    ' ============================================================
    '  3. VBA BUILT-IN OBJECTS & CONSTANTS
    ' ============================================================
    ' Constants
    ' MsgBox constants
    ' Color constants
    ' Comparison constants
    ' Date constants
    ' File attribute constants
    ' VarType constants
    Public ReadOnly VbaBuiltInObjects As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
        "Err", "Debug", "App"}

    Public ReadOnly VbaBuiltInConstants As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
        "vbNullString", "vbNullChar",
        "vbCr", "vbLf", "vbCrLf", "vbNewLine", "vbTab",
        "vbBack", "vbFormFeed", "vbVerticalTab",
        "vbOKOnly", "vbOKCancel", "vbAbortRetryIgnore", "vbYesNoCancel", "vbYesNo", "vbRetryCancel",
        "vbCritical", "vbQuestion", "vbExclamation", "vbInformation",
        "vbDefaultButton1", "vbDefaultButton2", "vbDefaultButton3", "vbDefaultButton4",
        "vbApplicationModal", "vbSystemModal",
        "vbOK", "vbCancel", "vbAbort", "vbRetry", "vbIgnore", "vbYes", "vbNo",
        "vbBlack", "vbRed", "vbGreen", "vbYellow", "vbBlue", "vbMagenta", "vbCyan", "vbWhite",
        "vbBinaryCompare", "vbTextCompare", "vbDatabaseCompare",
        "vbSunday", "vbMonday", "vbTuesday", "vbWednesday", "vbThursday", "vbFriday", "vbSaturday",
        "vbFirstJan1", "vbFirstFourDays", "vbFirstFullWeek",
        "vbUseSystem", "vbUseSystemDayOfWeek",
        "vbNormal", "vbReadOnly", "vbHidden", "vbSystem", "vbVolume", "vbDirectory", "vbArchive",
        "vbEmpty", "vbNull", "vbInteger", "vbLong", "vbSingle", "vbDouble", "vbCurrency",
        "vbDate", "vbString", "vbObject", "vbError", "vbBoolean", "vbVariant",
        "vbDataObject", "vbDecimal", "vbByte", "vbArray"
    }

    ' ============================================================
    '  4. ACCESS BUILT-IN OBJECTS & FUNCTIONS
    ' ============================================================
    ' Global objects
    ' Forms/Reports shortcuts
    Public ReadOnly AccessBuiltInObjects As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
       "DoCmd", "CurrentDb", "CurrentProject", "CodeDb", "CodeProject",
        "Screen", "Application", "DBEngine",
        "Modules", "References", "VBE",
        "Forms", "Reports",
        "Me", "Parent"
    }

    ' Domain aggregate functions
    ' SQL aggregate
    ' Conversion/Formatting
    ' Financial (duplicate cu VBA, dar păstrăm)
    Public ReadOnly AccessBuiltInFunctions As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
        "DLookup", "DSum", "DAvg", "DCount", "DMin", "DMax",
        "DFirst", "DLast", "DStDev", "DStDevP", "DVar", "DVarP",
        "Sum", "Avg", "Count", "Min", "Max", "First", "Last",
        "StDev", "StDevP", "Var", "VarP",
        "Nz", "Eval",
        "PMT", "FV", "PV", "NPER", "RATE", "IPMT", "PPMT"
    }

    ' ============================================================
    '  5. EXTERNAL LIBRARY PREFIXES (Early Binding)
    ' ============================================================
    ' ADO (ActiveX Data Objects)
    ' Scripting
    ' Office automation
    ' DAO (Data Access Objects)        
    ' Windows API structures        
    ' Other common
    Public ReadOnly ExternalLibraryPrefixes As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
        "ADODB", "ADOX", "ADOMD", "JRO",
        "Scripting",
        "Excel", "Word", "Outlook", "PowerPoint", "Access",
        "DAO",
        "VBA", "VB",
        "MSXML2", "Shell32", "WScript", "WshShell",
        "Regex", "RegExp"
    }

    ' ============================================================
    '  6. COMMON EXTERNAL CLASSES (Late Binding tracking)
    ' ============================================================
    Public ReadOnly ExternalClassNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
        "Scripting.FileSystemObject",
        "Scripting.Dictionary",
        "ADODB.Connection",
        "ADODB.Recordset",
        "ADODB.Command",
        "ADODB.Stream",
        "Excel.Application",
        "Word.Application",
        "Outlook.Application",
        "Shell.Application",
        "WScript.Shell",
        "MSXML2.DOMDocument",
        "VBScript.RegExp"
    }

    ' ============================================================
    '  7. VBA LANGUAGE KEYWORDS
    ' ============================================================
    ' --- Access modifiers / declarations ---
    ' --- Control flow ---
    ' --- Object and variable handling ---
    ' --- Logical and comparison operators ---
    ' --- Type conversion / casting ---
    ' --- Constants / values ---
    ' --- Misc language elements ---
    ReadOnly VbaKeywords As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
        "Public", "Private", "Friend", "Static", "Global", "Dim", "Const", "WithEvents",
        "Declare", "PtrSafe", "Lib", "Alias",
        "Type", "End", "Enum", "Property", "Get", "Let", "Set",
        "Sub", "Function", "Event", "Option", "Explicit", "Base", "Compare",
        "If", "Then", "Else", "ElseIf", "End If",
        "Select", "Case", "End Select",
        "Do", "Loop", "While", "Until",
        "For", "Each", "To", "Step", "Next", "Exit", "Continue",
        "GoTo", "On", "Error", "Resume",
        "With", "End With",
        "Stop", "Return", "GoSub",
        "New", "Set", "Nothing", "Is", "As", "ByVal", "ByRef", "Optional", "ParamArray",
        "And", "Or", "Not", "Xor", "Eqv", "Imp",
        "Like", "Mod", "Is", "In", "Of",
        "TypeOf",
        "True", "False", "Me",
        "Call", "ReDim", "Preserve", "Erase", "Option", "Attribute", "Implements",
        "RaiseEvent", "Handles", "Line", "Input", "Output", "Print", "Open", "Close",
        "Lock", "Unlock", "Put", "Get", "Seek", "Name", "Kill", "Width", "Write",
        "Random", "Binary", "Append",
        "ByRef", "ByVal", "Optional", "ParamArray", "Exit", "GoTo"
    }


    ' ============================================================
    '  HELPER FUNCTIONS
    ' ============================================================

    ''' <summary>
    ''' Verifică dacă un token string este built-in VBA/Access
    ''' </summary>
    Public Function IsBuiltIn(tokenString As String) As Boolean
        Return VbaBuiltInTypes.Contains(tokenString) OrElse
               VbaBuiltInFunctions.Contains(tokenString) OrElse
               VbaBuiltInObjects.Contains(tokenString) OrElse
               AccessBuiltInObjects.Contains(tokenString) OrElse
               AccessBuiltInFunctions.Contains(tokenString)
    End Function

    ''' <summary>
    ''' Verifică dacă un token are prefix de external library
    ''' </summary>
    Public Function IsExternalLibraryPrefix(tokenString As String) As Boolean
        For Each prefix In ExternalLibraryPrefixes
            If tokenString.StartsWith(prefix & ".", StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If
        Next
        Return False
    End Function

    ''' <summary>
    ''' Determină categoria de ignore pentru un token
    ''' </summary>
    Public Function GetIgnoreCategory(tokenString As String) As String
        If VbaBuiltInTypes.Contains(tokenString) Then Return "VBA Type"
        If VbaBuiltInFunctions.Contains(tokenString) Then Return "VBA Function"
        If VbaBuiltInObjects.Contains(tokenString) Then Return "VBA Object"
        If VbaBuiltInConstants.Contains(tokenString) Then Return "VBA Constant"
        If AccessBuiltInObjects.Contains(tokenString) Then Return "Access Object"
        If AccessBuiltInFunctions.Contains(tokenString) Then Return "Access Function"
        If IsExternalLibraryPrefix(tokenString) Then Return "External Library"
        If VbaKeywords.Contains(tokenString) Then Return "VBA Keyword"
        Return "Unknown"
    End Function

End Module