
Imports System.Reflection
Imports System.Text

Public Module DebugExtensionsV2
    ' Clasa pentru noduri de debug tree
    Public Class DebugNode
        Public Property Name As String
        Public Property Value As String
        Public Property DisplayText As String
        Public Property NodeType As String  ' "property", "object", "array", "primitive"
        Public Property OriginalObject As Object
        Public Property Children As New List(Of DebugNode)
        Public Property Count As Integer  ' pentru collections

        Public Overrides Function ToString() As String
            Return DisplayText
        End Function
    End Class


    <Runtime.CompilerServices.Extension()>
    Public Function DumpJson(obj As Object, Optional maxDepth As Integer = 2, Optional excludeBackingFields As Boolean = True) As String
        Dim visited As New HashSet(Of Object)(New ReferenceEqualityComparer())
        Return DumpJsonInternal(obj, 0, maxDepth, visited, Nothing, excludeBackingFields)
    End Function

    Private Function DumpJsonInternal(obj As Object, currentDepth As Integer, maxDepth As Integer,
                                      visited As HashSet(Of Object), parentObj As Object, excludeBackingFields As Boolean) As String
        ' Null check
        If obj Is Nothing Then Return "null"

        Dim objType = obj.GetType()

        ' === Primitive types & Strings ===
        If objType.IsPrimitive OrElse objType Is GetType(String) OrElse objType Is GetType(Decimal) Then
            Return FormatPrimitive(obj)
        End If

        ' === Check circular reference ===
        If objType.IsClass AndAlso Not TypeOf obj Is String Then
            If visited.Contains(obj) Then
                Dim identifier = GetObjectIdentifier(obj)
                Return $"""↻ {objType.Name}#{identifier}"""
            End If
            visited.Add(obj)
        End If

        ' === Max depth reached - show type info ===
        If currentDepth >= maxDepth Then
            Return FormatTypeInfo(obj, objType)
        End If

        ' === Collections: Dictionary ===
        If TypeOf obj Is IDictionary Then
            Return DumpDictionary(DirectCast(obj, IDictionary), currentDepth, maxDepth, visited, excludeBackingFields)
        End If

        ' === Collections: List/IEnumerable ===
        If TypeOf obj Is IEnumerable AndAlso Not TypeOf obj Is String Then
            Return DumpEnumerable(DirectCast(obj, IEnumerable), currentDepth, maxDepth, visited, excludeBackingFields)
        End If

        ' === Complex Object ===
        Return DumpComplexObject(obj, currentDepth, maxDepth, visited, parentObj, excludeBackingFields)
    End Function

    Private Function DumpComplexObject(obj As Object, currentDepth As Integer, maxDepth As Integer,
                                       visited As HashSet(Of Object), parentObj As Object, excludeBackingFields As Boolean) As String
        Dim sb As New StringBuilder()
        sb.Append("{")

        Dim objType = obj.GetType()
        Dim hasContent As Boolean = False

        ' Adaug tip la început pentru context
        If currentDepth > 0 Then
            sb.Append($"""_type"": ""{objType.Name}""")
            hasContent = True
        End If

        ' Properties
        For Each prop In objType.GetProperties(BindingFlags.Public Or BindingFlags.Instance)
            If Not prop.CanRead Then Continue For

            ' Skip backing fields (proprietăți care încep cu "_")
            If excludeBackingFields AndAlso prop.Name.StartsWith("_") Then Continue For

            Try
                Dim value = prop.GetValue(obj)

                ' Skip default values
                If IsDefaultValue(value, prop.PropertyType) Then Continue For

                ' Skip parent reference to avoid infinite loop
                If value IsNot Nothing AndAlso Object.ReferenceEquals(value, parentObj) Then
                    If hasContent Then sb.Append(",")
                    sb.Append($"""{prop.Name}"": ""⬆ PARENT""")
                    hasContent = True
                    Continue For
                End If

                ' Dump value
                Dim valueDump = DumpJsonInternal(value, currentDepth + 1, maxDepth, visited, obj, excludeBackingFields)

                If hasContent Then sb.Append(",")
                sb.Append($"""{prop.Name}"": {valueDump}")
                hasContent = True

            Catch ex As Exception
                ' Skip properties that throw exceptions
            End Try
        Next

        ' Fields (doar dacă nu sunt backing fields)
        For Each field In objType.GetFields(BindingFlags.Instance Or BindingFlags.NonPublic Or BindingFlags.Public)
            If field.IsSpecialName Then Continue For
            If excludeBackingFields AndAlso field.Name.StartsWith("_") Then Continue For

            Try
                Dim value = field.GetValue(obj)

                ' Skip default values
                If IsDefaultValue(value, field.FieldType) Then Continue For

                ' Skip parent reference
                If value IsNot Nothing AndAlso Object.ReferenceEquals(value, parentObj) Then
                    If hasContent Then sb.Append(",")
                    sb.Append($"""{field.Name}"": ""⬆ PARENT""")
                    hasContent = True
                    Continue For
                End If

                Dim valueDump = DumpJsonInternal(value, currentDepth + 1, maxDepth, visited, obj, excludeBackingFields)

                If hasContent Then sb.Append(",")
                sb.Append($"""{field.Name}"": {valueDump}")
                hasContent = True

            Catch ex As Exception
                ' Skip fields that throw exceptions
            End Try
        Next

        sb.Append("}")
        Return sb.ToString()
    End Function

    Private Function DumpDictionary(dict As IDictionary, currentDepth As Integer, maxDepth As Integer,
                                    visited As HashSet(Of Object), excludeBackingFields As Boolean) As String
        If dict.Count = 0 Then Return Nothing ' Skip empty

        ' La nivel maxim, arăt doar info
        If currentDepth >= maxDepth Then
            Return $"""📦 {dict.GetType().Name} ({dict.Count} items)"""
        End If

        Dim sb As New StringBuilder()
        sb.Append("{")

        Dim first As Boolean = True
        Dim maxItems As Integer = 10 ' Limitez la 10 items pentru lizibilitate
        Dim itemCount As Integer = 0

        For Each key In dict.Keys
            If itemCount >= maxItems Then
                If Not first Then sb.Append(",")
                sb.Append($"""..."": ""... și încă {dict.Count - maxItems} items""")
                Exit For
            End If

            Dim value = dict(key)

            If Not first Then sb.Append(",")
            first = False

            Dim keyStr = If(key Is Nothing, "null", key.ToString())
            Dim valueStr = DumpJsonInternal(value, currentDepth + 1, maxDepth, visited, Nothing, excludeBackingFields)

            sb.Append($"""{EscapeJsonString(keyStr)}"": {valueStr}")
            itemCount += 1
        Next

        sb.Append("}")
        Return sb.ToString()
    End Function

    Private Function DumpEnumerable(enumerable As IEnumerable, currentDepth As Integer, maxDepth As Integer,
                                    visited As HashSet(Of Object), excludeBackingFields As Boolean) As String
        Dim items As New List(Of Object)
        For Each item In enumerable
            items.Add(item)
        Next

        If items.Count = 0 Then Return Nothing ' Skip empty

        ' La nivel maxim, arăt doar info
        If currentDepth >= maxDepth Then
            Return $"""📦 {enumerable.GetType().Name} ({items.Count} items)"""
        End If

        Dim sb As New StringBuilder()
        sb.Append("[")

        Dim maxItems As Integer = 5 ' Limitez la 5 items pentru liste
        For i As Integer = 0 To Math.Min(items.Count - 1, maxItems - 1)
            If i > 0 Then sb.Append(",")
            sb.Append(DumpJsonInternal(items(i), currentDepth + 1, maxDepth, visited, Nothing, excludeBackingFields))
        Next

        If items.Count > maxItems Then
            sb.Append($", ""... +{items.Count - maxItems} more""")
        End If

        sb.Append("]")
        Return sb.ToString()
    End Function

    Private Function FormatPrimitive(obj As Object) As String
        If obj Is Nothing Then Return "null"

        Select Case obj.GetType()
            Case GetType(String)
                Dim str = obj.ToString()
                If str.Length > 50 Then str = str.Substring(0, 47) & "..."
                Return $"""{EscapeJsonString(str)}"""
            Case GetType(Boolean)
                Return If(DirectCast(obj, Boolean), "true", "false")
            Case GetType(DateTime)
                Return $"""{DirectCast(obj, DateTime).ToString("yyyy-MM-dd HH:mm:ss")}"""
            Case Else
                Return obj.ToString()
        End Select
    End Function

    Private Function FormatTypeInfo(obj As Object, objType As Type) As String
        Dim typeName = objType.Name

        ' Pentru colecții
        If TypeOf obj Is ICollection Then
            Dim count = DirectCast(obj, ICollection).Count
            Return $"""📦 {typeName} ({count} items)"""
        End If

        ' Pentru obiecte custom
        Dim identifier = GetObjectIdentifier(obj)
        Return $"""🔗 {typeName}#{identifier}"""
    End Function

    Private Function GetObjectIdentifier(obj As Object) As String
        If obj Is Nothing Then Return "null"

        Dim objType = obj.GetType()

        ' Încerc să găsesc o proprietate Name
        Dim nameProp = objType.GetProperty("Name", BindingFlags.Public Or BindingFlags.Instance)
        If nameProp IsNot Nothing AndAlso nameProp.CanRead Then
            Try
                Dim nameValue = nameProp.GetValue(obj)
                If nameValue IsNot Nothing AndAlso Not String.IsNullOrEmpty(nameValue.ToString()) Then
                    Return nameValue.ToString()
                End If
            Catch
            End Try
        End If

        ' Altfel, folosesc doar tipul (mai lizibil decât hashcode)
        Return objType.Name
    End Function

    Private Function IsDefaultValue(value As Object, valueType As Type) As Boolean
        ' Null/Nothing
        If value Is Nothing Then Return True

        ' String gol
        If TypeOf value Is String AndAlso String.IsNullOrEmpty(DirectCast(value, String)) Then Return True

        ' Colecții goale
        If TypeOf value Is ICollection AndAlso DirectCast(value, ICollection).Count = 0 Then Return True

        ' Value types cu valoare default
        If valueType.IsValueType Then
            Dim defaultValue = Activator.CreateInstance(valueType)
            Return value.Equals(defaultValue)
        End If

        Return False
    End Function

    Private Function EscapeJsonString(str As String) As String
        If String.IsNullOrEmpty(str) Then Return ""

        ' NU mai escapăm ghilimelele - le lăsăm așa cum sunt
        Return str.Replace("\", "\\") _
              .Replace(vbCr, "\r") _
              .Replace(vbLf, "\n") _
              .Replace(vbTab, "\t")
    End Function

    ' Helper class for reference equality comparison
    Private Class ReferenceEqualityComparer
        Inherits EqualityComparer(Of Object)

        Public Overrides Function Equals(x As Object, y As Object) As Boolean
            Return Object.ReferenceEquals(x, y)
        End Function

        Public Overrides Function GetHashCode(obj As Object) As Integer
            Return Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj)
        End Function
    End Class
End Module