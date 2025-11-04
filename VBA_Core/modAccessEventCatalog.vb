' modAccessEventCatalog.vb
Imports System.Collections.Generic

Public Module modAccessEventCatalog

    ' ====================================================
    '  EVENIMENTE SPECIFICE PE TIPURI DE OBIECTE ACCESS
    ' ====================================================

    Private ReadOnly FormEvents As String() = {
        "OnOpen", "OnLoad", "OnResize", "OnActivate", "OnDeactivate",
        "OnCurrent", "BeforeInsert", "AfterInsert", "BeforeUpdate",
        "AfterUpdate", "BeforeDelConfirm", "AfterDelConfirm",
        "OnClose", "OnUnload", "OnError", "OnApplyFilter", "OnFilter",
        "OnNotInList", "OnDirty", "OnUndo", "OnPaint", "OnChangeView",
        "OnTimer", "AfterScroll", "OnRequery"
    }

    Private ReadOnly ReportEvents As String() = {
        "OnOpen", "OnLoad", "OnActivate", "OnDeactivate", "OnClose",
        "OnNoData", "OnError", "OnPage", "OnFormat", "OnPrint", "OnRetreat",
        "OnChange", "OnMove", "OnResize", "AfterScroll", "BeforeUpdate",
        "AfterUpdate"
    }

    Private ReadOnly ControlEventsCommon As String() = {
        "OnClick", "OnDblClick", "OnMouseDown", "OnMouseUp", "OnMouseMove",
        "OnEnter", "OnExit", "OnGotFocus", "OnLostFocus",
        "OnKeyDown", "OnKeyUp", "OnKeyPress", "OnChange", "OnBeforeUpdate",
        "OnAfterUpdate", "OnDirty", "OnValidate", "OnMouseWheel", "OnLostFocus"
    }

    Private ReadOnly ControlEventsInput As String() = {
        "OnBeforeUpdate", "OnAfterUpdate", "OnChange", "OnDirty", "OnValidate",
        "OnNotInList"
    }

    Private ReadOnly ControlEventsButton As String() = {
        "OnClick", "OnDblClick", "OnMouseDown", "OnMouseUp", "OnMouseMove",
        "OnEnter", "OnExit"
    }

    Private ReadOnly ControlEventsSubForm As String() = {
        "OnEnter", "OnExit", "OnGotFocus", "OnLostFocus",
        "OnClick", "OnDblClick", "OnMouseDown", "OnMouseUp", "OnMouseMove",
        "OnCurrent", "OnLoad", "OnOpen", "OnClose", "OnError", "AfterScroll", "OnRequery"
    }

    Private ReadOnly RecordsetEvents As String() = {
        "BeforeInsert", "AfterInsert", "BeforeUpdate", "AfterUpdate",
        "BeforeDelete", "AfterDelete", "AfterScroll", "OnFetch", "OnRequery",
        "OnChange", "OnMove"
    }

    ' ===============================================
    '  PUBLIC API
    ' ===============================================
    ' =======================================================================================
    '  Identificare EVENT HANDLER din nume de metodă: Form_Click, TextBox1_AfterUpdate etc.
    ' =======================================================================================
    Public Function TryParseEventHandlerName(methodName As String,
                                         Optional ByRef objectName As String = Nothing,
                                         Optional ByRef eventName As String = Nothing,
                                         Optional ByRef objectType As String = Nothing) As Boolean
        If String.IsNullOrEmpty(methodName) Then Return False

        ' tipice: Form_Click / Report_Open / Command0_Click / Text1_Change
        Dim parts = methodName.Split("_"c)
        If parts.Length < 2 Then Return False

        objectName = parts(0)
        eventName = String.Join("_", parts.Skip(1))

        ' 1️⃣ Dacă e "Form" sau "Report"
        If objectName.Equals("Form", StringComparison.OrdinalIgnoreCase) Then
            objectType = "Form"
            Return TryIsKnownEvent("Form", "On" & eventName)
        End If
        If objectName.Equals("Report", StringComparison.OrdinalIgnoreCase) Then
            objectType = "Report"
            Return TryIsKnownEvent("Report", "On" & eventName)
        End If

        ' 2️⃣ Poate fi control (TextBox, ComboBox, CommandButton etc.)
        ' Eliminăm eventualul index numeric (Text1 -> TextBox, Command0 -> CommandButton)
        Dim ctrlBase As String = ""
        If objectName.StartsWith("Text", StringComparison.OrdinalIgnoreCase) Then ctrlBase = "TextBox"
        If objectName.StartsWith("Combo", StringComparison.OrdinalIgnoreCase) Then ctrlBase = "ComboBox"
        If objectName.StartsWith("Command", StringComparison.OrdinalIgnoreCase) Then ctrlBase = "CommandButton"
        If objectName.StartsWith("List", StringComparison.OrdinalIgnoreCase) Then ctrlBase = "ListBox"
        If objectName.StartsWith("Check", StringComparison.OrdinalIgnoreCase) Then ctrlBase = "CheckBox"
        If objectName.StartsWith("Option", StringComparison.OrdinalIgnoreCase) Then ctrlBase = "OptionButton"
        If objectName.StartsWith("Toggle", StringComparison.OrdinalIgnoreCase) Then ctrlBase = "ToggleButton"
        If objectName.StartsWith("SubForm", StringComparison.OrdinalIgnoreCase) Then ctrlBase = "SubForm"

        If ctrlBase <> "" Then
            objectType = ctrlBase
            Return TryIsKnownEvent(ctrlBase, "On" & eventName)
        End If

        Return False
    End Function

    Public Function GetEventsForObjectType(objType As String) As List(Of String)
        Dim eventsSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        Select Case True
            Case objType.Equals("Form", StringComparison.OrdinalIgnoreCase)
                eventsSet.UnionWith(FormEvents)

            Case objType.Equals("Report", StringComparison.OrdinalIgnoreCase)
                eventsSet.UnionWith(ReportEvents)

            Case objType.Equals("SubForm", StringComparison.OrdinalIgnoreCase)
                eventsSet.UnionWith(ControlEventsSubForm)

            Case objType.Equals("CommandButton", StringComparison.OrdinalIgnoreCase)
                eventsSet.UnionWith(ControlEventsButton)

            Case objType.Equals("TextBox", StringComparison.OrdinalIgnoreCase),
                 objType.Equals("ComboBox", StringComparison.OrdinalIgnoreCase),
                 objType.Equals("ListBox", StringComparison.OrdinalIgnoreCase),
                 objType.Equals("OptionButton", StringComparison.OrdinalIgnoreCase),
                 objType.Equals("CheckBox", StringComparison.OrdinalIgnoreCase),
                 objType.Equals("ToggleButton", StringComparison.OrdinalIgnoreCase)
                eventsSet.UnionWith(ControlEventsCommon)
                eventsSet.UnionWith(ControlEventsInput)

            Case objType.Equals("Label", StringComparison.OrdinalIgnoreCase),
                 objType.Equals("Image", StringComparison.OrdinalIgnoreCase),
                 objType.Equals("Line", StringComparison.OrdinalIgnoreCase),
                 objType.Equals("Rectangle", StringComparison.OrdinalIgnoreCase),
                 objType.Equals("Frame", StringComparison.OrdinalIgnoreCase),
                 objType.Equals("TabControl", StringComparison.OrdinalIgnoreCase)
                eventsSet.UnionWith(ControlEventsCommon)

            Case objType.Equals("Recordset", StringComparison.OrdinalIgnoreCase)
                eventsSet.UnionWith(RecordsetEvents)

            Case Else
                ' fallback: toate evenimentele comune
                eventsSet.UnionWith(ControlEventsCommon)
        End Select

        Return eventsSet.ToList()
    End Function

    Public Function TryIsKnownEvent(objType As String, eventName As String) As Boolean
        Return GetEventsForObjectType(objType).Any(Function(e) e.Equals(eventName, StringComparison.OrdinalIgnoreCase))
    End Function

    Public Function GetAllCommonEvents() As List(Of String)
        Dim all As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        all.UnionWith(ControlEventsCommon)
        all.UnionWith(ControlEventsInput)
        Return all.ToList()
    End Function

End Module
