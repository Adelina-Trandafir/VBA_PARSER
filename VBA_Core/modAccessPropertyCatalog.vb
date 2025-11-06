Imports System.Collections.Generic
Imports System.Linq

Public Module modAccessPropertyCatalog

    ' ===========================
    ' Common control properties
    ' ===========================
    Private ReadOnly CommonCtrlProps As String() = {
        "Name", "Tag", "ControlTipText", "HelpContextId", "StatusBarText",
        "Visible", "Enabled", "Locked", "TabIndex", "TabStop", "TabFixedHeight", "TabFixedWidth",
        "Left", "Top", "Width", "Height", "ColumnOrder", "Layout", "PaddingTop", "PaddingBottom", "PaddingLeft", "PaddingRight",
        "ForeColor", "BackColor", "BackStyle", "SpecialEffect", "BorderStyle", "BorderColor", "BorderWidth",
        "FontName", "FontSize", "FontWeight", "FontBold", "FontItalic", "FontUnderline", "TextAlign"
    }

    ' ======================================
    ' Control-type specific properties map
    ' ======================================
    Private ReadOnly CtrlProps As New Dictionary(Of String, String())(StringComparer.OrdinalIgnoreCase) From {
        {"TextBox", {
            "ControlSource", "DefaultValue", "Format", "DecimalPlaces", "InputMask",
            "Value", "Text", "EnterKeyBehavior", "SelectionMargin", "WordWrap", "ScrollBars",
            "AllowAutoCorrect", "ImeMode", "ImeSentenceMode", "AutoTab", "AutoExpand"
        }},
        {"ComboBox", {
            "ControlSource", "RowSource", "RowSourceType", "BoundColumn", "ColumnCount", "ColumnHeads",
            "ColumnWidths", "LimitToList", "ListRows", "ListWidth", "AutoExpand", "Value"
        }},
        {"ListBox", {
            "RowSource", "RowSourceType", "ColumnCount", "ColumnHeads", "ColumnWidths",
            "BoundColumn", "MultiSelect", "Value"
        }},
        {"Label", {
            "Caption", "HyperlinkAddress", "HyperlinkSubAddress"
        }},
        {"CommandButton", {
            "Caption", "Picture", "PictureType", "PictureCaptionArrangement", "TakeFocusOnClick",
            "OnClick", "OnDblClick"
        }},
        {"OptionButton", {
            "Caption", "OptionValue", "Value"
        }},
        {"CheckBox", {
            "Caption", "TripleState", "Value"
        }},
        {"ToggleButton", {
            "Caption", "Value"
        }},
        {"SubForm", {
            "SourceObject", "LinkChildFields", "LinkMasterFields", "Class", "Format", "FastLaserPrinting",
            "OnEnter", "OnExit", "OnGotFocus", "OnLostFocus"
        }},
        {"Frame", {
            "Caption", "SpecialEffect"
        }},
        {"TabControl", {
            "Value", "Style", "MultiRow"
        }},
        {"Image", {
            "Picture", "PictureType", "PictureAlignment", "PictureTiling", "PictureSizeMode"
        }},
        {"Line", {
            "BorderStyle", "BorderWidth", "BorderColor"
        }},
        {"Rectangle", {
            "BackStyle", "BorderStyle", "BorderWidth", "BorderColor", "SpecialEffect"
        }}
    }

    ' ======================================
    ' Form properties (container)
    ' ======================================
    Private ReadOnly FormProps As String() = {
        "Caption", "RecordSource", "Filter", "FilterOn", "OrderBy", "OrderByOn",
        "AllowEdits", "AllowAdditions", "AllowDeletions", "AllowFilters", "DataEntry",
        "Cycle", "DefaultView", "ViewsAllowed", "AllowFormView", "AllowDatasheetView", "AllowPivotTableView", "AllowPivotChartView",
        "NavigationButtons", "DividingLines", "ScrollBars", "AutoResize", "AutoCenter", "Picture", "PictureType",
        "BorderStyle", "PopUp", "Modal", "MinMaxButtons", "CloseButton", "ControlBox", "Moveable",
        "Width", "InsideWidth", "InsideHeight", "Section(acDetail)", "Section(acHeader)", "Section(acFooter)",
        "OnCurrent", "OnOpen", "OnClose", "OnLoad", "OnUnload", "OnResize", "OnTimer", "TimerInterval",
        "OnApplyFilter", "OnDirty", "OnNotInList", "OnFilter", "OnError", "OnActivate", "OnDeactivate"
    }

    ' ==================================================
    ' Subform control “forwarded” form properties (read)
    ' ==================================================
    Private ReadOnly SubformForwarded As String() = {
        "Form", "Form.Caption", "Form.RecordSource", "Form.Filter", "Form.FilterOn",
        "Form.AllowEdits", "Form.AllowAdditions", "Form.AllowDeletions", "Form.AllowFilters", "Form.DataEntry",
        "Form.OnCurrent", "Form.OnOpen", "Form.OnClose", "Form.OnLoad", "Form.OnError"
    }

    ' ==================================================
    ' Events as resolvable property-like names
    ' ==================================================
    Private ReadOnly CommonEvents As String() = {
        "OnClick", "OnDblClick", "OnMouseDown", "OnMouseUp", "OnMouseMove",
        "OnKeyDown", "OnKeyUp", "OnKeyPress", "OnEnter", "OnExit", "OnGotFocus", "OnLostFocus",
        "AfterUpdate", "BeforeUpdate", "OnChange"
    }

    ' ==================================================
    ' Control-type specific event map
    ' ==================================================
    Private ReadOnly CtrlEvents As New Dictionary(Of String, String())(StringComparer.OrdinalIgnoreCase) From {
        {"CommandButton", {"Click", "DblClick"}},
        {"TextBox", {"AfterUpdate", "BeforeUpdate", "Change", "GotFocus", "LostFocus"}},
        {"ComboBox", {"AfterUpdate", "BeforeUpdate", "Change", "Enter", "Exit"}},
        {"ListBox", {"AfterUpdate", "BeforeUpdate", "Enter", "Exit"}},
        {"CheckBox", {"Click", "AfterUpdate"}},
        {"OptionButton", {"Click", "AfterUpdate"}},
        {"ToggleButton", {"Click", "AfterUpdate"}},
        {"Form", {"Load", "Open", "Resize", "Unload", "Close", "Activate", "Deactivate"}},
        {"Report", {"Open", "Close", "Activate", "Deactivate"}}
    }

    ' ===========================
    ' Public API
    ' ===========================

    Public Function GetControlProperties(ctrlType As String) As List(Of String)
        Dim setAll As New HashSet(Of String)(CommonCtrlProps, StringComparer.OrdinalIgnoreCase)

        ' Include common events pentru toate controalele
        For Each eName In CommonEvents
            setAll.Add(eName)
        Next

        If CtrlProps.ContainsKey(ctrlType) Then
            For Each p In CtrlProps(ctrlType)
                setAll.Add(p)
            Next
        End If

        ' SubForm aliasuri: Subform, SubForm, SubformControl
        If ctrlType.Equals("Subform", StringComparison.OrdinalIgnoreCase) _
           OrElse ctrlType.Equals("SubForm", StringComparison.OrdinalIgnoreCase) _
           OrElse ctrlType.Equals("SubformControl", StringComparison.OrdinalIgnoreCase) Then
            For Each p In SubformForwarded
                setAll.Add(p)
            Next
        End If

        Return setAll.ToList()
    End Function

    Public Function GetFormProperties() As List(Of String)
        Dim setAll As New HashSet(Of String)(FormProps, StringComparer.OrdinalIgnoreCase)
        For Each eName In CommonEvents
            setAll.Add(eName)
        Next
        Return setAll.ToList()
    End Function

    ' ========= Extensibility =========
    Public Sub AddControlProperty(ctrlType As String, propName As String)
        If Not CtrlProps.ContainsKey(ctrlType) Then
            CtrlProps(ctrlType) = Array.Empty(Of String)()
        End If
        Dim l = CtrlProps(ctrlType).ToList()
        If Not l.Any(Function(x) x.Equals(propName, StringComparison.OrdinalIgnoreCase)) Then
            l.Add(propName)
            CtrlProps(ctrlType) = l.ToArray()
        End If
    End Sub

    Public Sub AddFormProperty(propName As String)
        If Not FormProps.Contains(propName, StringComparer.OrdinalIgnoreCase) Then
            Dim l = FormProps.ToList()
            l.Add(propName)
            Dim arr = l.ToArray()
            ReplaceFormProps(arr)
        End If
    End Sub

    Private Sub ReplaceFormProps(newArr As String())
        Dim fi = GetType(modAccessPropertyCatalog).GetField("FormProps", Reflection.BindingFlags.NonPublic Or Reflection.BindingFlags.Static)
        fi.SetValue(Nothing, newArr)
    End Sub

    ' ========= Helpers pentru resolver =========

    ' Întoarce True dacă numele e proprietate validă pt. tipul de control
    Public Function TryIsKnownControlProperty(ctrlType As String, propName As String) As Boolean
        Return GetControlProperties(ctrlType).Any(Function(p) p.Equals(propName, StringComparison.OrdinalIgnoreCase))
    End Function

    ' Întoarce True dacă numele e proprietate validă pentru Form
    Public Function TryIsKnownFormProperty(propName As String) As Boolean
        Return GetFormProperties().Any(Function(p) String.Equals(p, propName, StringComparison.OrdinalIgnoreCase))
    End Function

    ' Întoarce True dacă evenimentul e cunoscut pentru un tip de control
    Public Function TryIsKnownControlEvent(ctrlType As String, evtName As String) As Boolean
        Return CtrlEvents.ContainsKey(ctrlType) AndAlso
               CtrlEvents(ctrlType).Any(Function(e) e.Equals(evtName, StringComparison.OrdinalIgnoreCase))
    End Function

    ' Întoarce True dacă evenimentul e cunoscut pentru Form sau Report
    Public Function TryIsKnownFormEvent(evtName As String) As Boolean
        Return CtrlEvents("Form").Contains(evtName, StringComparer.OrdinalIgnoreCase) OrElse
               CtrlEvents("Report").Contains(evtName, StringComparer.OrdinalIgnoreCase)
    End Function

    ' Încearcă să determine dacă un nume de membru e eveniment (ex. pentru legături OnClick = "=Macro()" sau „[Event Procedure]”)
    Public Function IsEventName(memberName As String) As Boolean
        Return CommonEvents.Any(Function(eName) eName.Equals(memberName, StringComparison.OrdinalIgnoreCase))
    End Function

End Module
