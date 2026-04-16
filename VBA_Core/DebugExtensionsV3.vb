Imports System.Text
Imports System.Reflection
Imports System.Collections

Public Module DebugExtensionsV3

    ' Clasa pentru noduri de debug tree
    Public Class DebugNodeV3
        Public Property Name As String
        Public Property Value As String
        Public Property NodeType As String  ' "property", "object", "array", "primitive", "collapsed"
        Public Property OriginalObject As Object  ' ÎNTOTDEAUNA păstrat, chiar și pentru collapsed
        Public Property Children As New List(Of DebugNodeV3)
        Public Property IsReference As Boolean
        Public Property LinkedObject As Object
        Public Property Count As Integer
        Public Property IsExpandable As Boolean  ' True dacă e un obiect complex care poate fi expandat
        Public ReadOnly Property DisplayText As String
            Get
                Dim prefix As String = ""
                If IsExpandable AndAlso NodeType = "collapsed" Then prefix = "🔍 "
                If NodeType = "reference" Then prefix = "↻ "
                If NodeType = "object" Then prefix = "🧩 "

                Dim val = If(Value, "")
                Return $"{prefix}{If(String.IsNullOrEmpty(Name), val, $"{Name}: {val}")}"
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return DisplayText
        End Function
    End Class

    Private ReadOnly TokenLikeTypes As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {"Token", "MethodLine", "ParamInfo"}

    <Runtime.CompilerServices.Extension()>
    Public Function DumpToTree(obj As Object, Optional maxDepth As Integer = 4, Optional excludeBackingFields As Boolean = True) As DebugNodeV3
        Dim visited As New HashSet(Of Object)(New ReferenceEqualityComparer())
        Return BuildDebugNode(obj, "Root", 0, maxDepth, visited, Nothing, excludeBackingFields)
    End Function

    Private Function BuildDebugNode(obj As Object, propertyName As String, currentDepth As Integer,
                                    maxDepth As Integer, visited As HashSet(Of Object),
                                    parentObj As Object, excludeBackingFields As Boolean) As DebugNodeV3

        Dim node As New DebugNodeV3() With {
            .Name = propertyName,
            .OriginalObject = obj
        }

        ' === NULL ===
        If obj Is Nothing Then
            node.Value = "null"
            'node.DisplayText = $"{propertyName}: null"
            node.NodeType = "primitive"
            node.IsExpandable = False
            Return node
        End If

        Dim objType = obj.GetType()

        ' === PRIMITIVES & STRINGS ===
        If objType.IsPrimitive OrElse objType Is GetType(String) OrElse objType Is GetType(Decimal) OrElse objType Is GetType(DateTime) Then
            node.Value = FormatPrimitiveValue(obj)
            'node.DisplayText = $"{propertyName}: {node.Value}"
            node.NodeType = "primitive"
            node.IsExpandable = False
            Return node
        End If

        ' === CIRCULAR REFERENCE ===
        If objType.IsClass AndAlso Not TypeOf obj Is String Then
            If visited.Contains(obj) Then
                Dim identifier = GetObjectIdentifier(obj)
                node.Value = $"↻ {objType.Name}#{identifier}"
                'node.DisplayText = $"{propertyName}: {node.Value}"
                node.NodeType = "reference"
                node.IsExpandable = False
                Return node
            End If
            visited.Add(obj)
        End If

        ' === MAX DEPTH - dar PĂSTRĂM OBIECTUL pentru expand ulterior ===
        If currentDepth >= maxDepth Then
            node.Value = GetTypeInfoString(obj, objType)
            'node.DisplayText = $"{propertyName}: {node.Value} 🔍"  ' 🔍 = click pentru expand
            node.NodeType = "collapsed"
            node.IsExpandable = True  ' IMPORTANT!
            Return node
        End If

        ' === DICTIONARY ===
        If TypeOf obj Is IDictionary Then
            Return BuildDictionaryNode(DirectCast(obj, IDictionary), propertyName, currentDepth, maxDepth, visited, excludeBackingFields, node)
        End If

        ' === ENUMERABLE (List, Array, etc) ===
        If TypeOf obj Is IEnumerable AndAlso Not TypeOf obj Is String Then
            Return BuildEnumerableNode(DirectCast(obj, IEnumerable), propertyName, currentDepth, maxDepth, visited, excludeBackingFields, node)
        End If

        ' === COMPLEX OBJECT ===
        If TokenLikeTypes.Contains(objType.Name) Then
            Return BuildTokenLikeNode(obj, propertyName, currentDepth, maxDepth, visited, node)
        End If

        Return BuildComplexObjectNode(obj, propertyName, currentDepth, maxDepth, visited, parentObj, excludeBackingFields, node)
    End Function

    Private Function BuildComplexObjectNode(obj As Object, propertyName As String, currentDepth As Integer,
                                        maxDepth As Integer, visited As HashSet(Of Object),
                                        parentObj As Object, excludeBackingFields As Boolean,
                                        node As DebugNodeV3) As DebugNodeV3
        Dim objType = obj.GetType()
        node.NodeType = "object"
        node.Value = objType.Name
        node.IsExpandable = True

        ' === PROPERTIES ===
        For Each prop In objType.GetProperties(BindingFlags.Public Or BindingFlags.Instance)
            If Not prop.CanRead Then Continue For
            If excludeBackingFields AndAlso prop.Name.StartsWith("_") Then Continue For

            Try
                Dim value = prop.GetValue(obj)
                If IsDefaultValue(value, prop.PropertyType) Then Continue For

                Dim childNode = BuildDebugNode(value, prop.Name, currentDepth + 1, maxDepth, visited, obj, excludeBackingFields)

                ' ============================================================
                ' 🔹 Detectăm dacă e o referință internă (Token, MethodInfo, etc.)
                ' ============================================================
                If value IsNot Nothing Then
                    Dim tName = value.GetType().Name.ToLowerInvariant()
                    If tName.Contains("token") OrElse
                   tName.Contains("methodinfo") OrElse
                   tName.Contains("paraminfo") OrElse
                   tName.Contains("propertyinfo") OrElse
                   tName.Contains("functioninfo") OrElse
                   tName.Contains("modulecontainer") Then
                        childNode.IsReference = True
                        childNode.LinkedObject = value
                    End If
                End If

                ' ============================================================
                ' 🔹 Caz special: ResolvedScope = "module_var"
                '    -> caută variabila reală în ParentModule.Variables
                ' ============================================================
                If prop.Name.Equals("ResolvedScope", StringComparison.OrdinalIgnoreCase) AndAlso
               TypeOf value Is String AndAlso
               CStr(value).Equals("module_var", StringComparison.OrdinalIgnoreCase) Then

                    ' Obținem ParentModule și ResolvedName din obiectul părinte
                    Dim resolvedNameProp = objType.GetProperty("ResolvedName")
                    Dim parentModuleProp = objType.GetProperty("ParentModule")

                    If resolvedNameProp IsNot Nothing AndAlso parentModuleProp IsNot Nothing Then
                        Dim resolvedName = TryCast(resolvedNameProp.GetValue(obj), String)
                        Dim parentModule = TryCast(parentModuleProp.GetValue(obj), ModuleContainer)

                        If Not String.IsNullOrEmpty(resolvedName) AndAlso parentModule IsNot Nothing Then
                            Dim foundVar = parentModule.Variables.FirstOrDefault(Function(v) v.Name.Equals(resolvedName, StringComparison.OrdinalIgnoreCase))
                            If foundVar IsNot Nothing Then
                                childNode.IsReference = True
                                childNode.LinkedObject = foundVar
                                childNode.Value &= $" → {foundVar.Type}"
                            End If
                        End If
                    End If
                End If

                node.Children.Add(childNode)

            Catch
                ' Ignorăm proprietăți problematice
            End Try
        Next

        ' === FIELDS (inclusiv private) ===
        For Each field In objType.GetFields(BindingFlags.Instance Or BindingFlags.NonPublic Or BindingFlags.Public)
            If field.IsSpecialName Then Continue For
            If excludeBackingFields AndAlso field.Name.StartsWith("_") Then Continue For

            Try
                Dim value = field.GetValue(obj)
                If IsDefaultValue(value, field.FieldType) Then Continue For

                ' Detectăm referință către obiectul părinte
                If value IsNot Nothing AndAlso Object.ReferenceEquals(value, parentObj) Then
                    Dim parentNode As New DebugNodeV3() With {
                    .Name = field.Name,
                    .Value = "⬆ PARENT",
                    .NodeType = "reference",
                    .OriginalObject = value,
                    .IsExpandable = False
                }
                    node.Children.Add(parentNode)
                    Continue For
                End If

                Dim childNode = BuildDebugNode(value, field.Name, currentDepth + 1, maxDepth, visited, obj, excludeBackingFields)

                ' Marchează intern dacă e tip relevant
                If value IsNot Nothing Then
                    Dim tName = value.GetType().Name.ToLowerInvariant()
                    If tName.Contains("token") OrElse
                   tName.Contains("methodinfo") OrElse
                   tName.Contains("paraminfo") OrElse
                   tName.Contains("propertyinfo") OrElse
                   tName.Contains("functioninfo") OrElse
                   tName.Contains("modulecontainer") Then
                        childNode.IsReference = True
                        childNode.LinkedObject = value
                    End If
                End If

                node.Children.Add(childNode)

            Catch
                ' Skip field
            End Try
        Next

        Return node
    End Function

    Private Function BuildTokenLikeNode(obj As Object, propertyName As String, currentDepth As Integer,
                                    maxDepth As Integer, visited As HashSet(Of Object),
                                    node As DebugNodeV3) As DebugNodeV3

        Dim objType = obj.GetType()
        node.NodeType = "object"
        'node.DisplayText = $"{propertyName}: {objType.Name}"
        node.Value = objType.Name
        node.IsExpandable = True

        ' === Enumerăm proprietăți relevante ===
        For Each prop In objType.GetProperties(BindingFlags.Public Or BindingFlags.Instance)
            If Not prop.CanRead Then Continue For

            ' Sărim peste referințe mari / potențial ciclice
            If {"Parent", "SourceLine", "WithContext", "ResolvedRef"}.Contains(prop.Name) Then Continue For

            Try
                Dim value = prop.GetValue(obj)
                ' 🔹 Ignoră NULL sau string gol
                If value Is Nothing Then Continue For
                If TypeOf value Is String AndAlso String.IsNullOrEmpty(DirectCast(value, String)) Then Continue For
                If TypeOf value Is Boolean AndAlso DirectCast(value, Boolean) = False Then Continue For

                ' simplificăm pentru valori null sau primitive
                Dim valStr As String
                If value Is Nothing Then
                    valStr = "null"
                ElseIf prop.PropertyType.IsPrimitive OrElse TypeOf value Is String OrElse value.GetType().IsEnum Then
                    valStr = value.ToString()
                ElseIf TypeOf value Is ICollection Then
                    valStr = $"[{DirectCast(value, ICollection).Count} items]"
                Else
                    valStr = $"[{value.GetType().Name}]"
                End If

                Dim childNode As New DebugNodeV3() With {
                .Name = prop.Name,
                .Value = valStr,
                .NodeType = "property",
                .IsExpandable = False
            } '                .DisplayText = $"{prop.Name}: {valStr}",


                node.Children.Add(childNode)

            Catch ex As Exception
                ' Ignoră proprietăți inaccesibile
            End Try
        Next

        Return node
    End Function

    Private Function BuildDictionaryNode(dict As IDictionary, propertyName As String, currentDepth As Integer,
                                         maxDepth As Integer, visited As HashSet(Of Object),
                                         excludeBackingFields As Boolean, node As DebugNodeV3) As DebugNodeV3
        node.NodeType = "dictionary"
        node.Count = dict.Count
        'node.DisplayText = $"{propertyName}: {dict.GetType().Name} ({dict.Count} items)"
        node.Value = $"{dict.GetType().Name} ({dict.Count})"
        node.IsExpandable = dict.Count > 0

        If dict.Count = 0 Then Return node

        Dim maxItems As Integer = 20
        Dim itemCount As Integer = 0

        For Each key In dict.Keys
            If itemCount >= maxItems Then
                Dim moreNode As New DebugNodeV3() With {
                    .Name = "...",
                    .Value = $"... și încă {dict.Count - maxItems} items",
                    .NodeType = "info",
                    .IsExpandable = False
                }
                '                    .DisplayText = $"... și încă {dict.Count - maxItems} items",

                node.Children.Add(moreNode)
                Exit For
            End If

            Dim value = dict(key)
            Dim keyStr = If(key Is Nothing, "null", key.ToString())

            Dim childNode = BuildDebugNode(value, $"[{keyStr}]", currentDepth + 1, maxDepth, visited, Nothing, excludeBackingFields)
            node.Children.Add(childNode)
            itemCount += 1
        Next

        Return node
    End Function

    Private Function BuildEnumerableNode(enumerable As IEnumerable, propertyName As String, currentDepth As Integer,
                                         maxDepth As Integer, visited As HashSet(Of Object),
                                         excludeBackingFields As Boolean, node As DebugNodeV3) As DebugNodeV3
        Dim items As New List(Of Object)
        For Each item In enumerable
            items.Add(item)
        Next

        node.NodeType = "array"
        node.Count = items.Count
        'node.DisplayText = $"{propertyName}: {enumerable.GetType().Name} ({items.Count} items)"
        node.Value = $"{enumerable.GetType().Name} ({items.Count})"
        node.IsExpandable = items.Count > 0

        If items.Count = 0 Then Return node

        Dim maxItems As Integer = 20
        For i As Integer = 0 To Math.Min(items.Count - 1, maxItems - 1)
            Dim childNode = BuildDebugNode(items(i), $"[{i}]", currentDepth + 1, maxDepth, visited, Nothing, excludeBackingFields)
            node.Children.Add(childNode)
        Next

        If items.Count > maxItems Then
            Dim moreNode As New DebugNodeV3() With {
                .Name = "...",
                .Value = $"... +{items.Count - maxItems} more",
                .NodeType = "info",
                .IsExpandable = False
            } '                .DisplayText = $"... +{items.Count - maxItems} more",

            node.Children.Add(moreNode)
        End If

        Return node
    End Function

    Private Function FormatPrimitiveValue(obj As Object) As String
        If obj Is Nothing Then Return "null"

        Select Case obj.GetType()
            Case GetType(String)
                Dim str = obj.ToString()
                If str.Length > 100 Then str = str.Substring(0, 97) & "..."
                Return str
            Case GetType(Boolean)
                Return If(DirectCast(obj, Boolean), "true", "false")
            Case GetType(DateTime)
                Return DirectCast(obj, DateTime).ToString("yyyy-MM-dd HH:mm:ss")
            Case Else
                Return obj.ToString()
        End Select
    End Function

    Private Function GetTypeInfoString(obj As Object, objType As Type) As String
        If TypeOf obj Is ICollection Then
            Dim count = DirectCast(obj, ICollection).Count
            Return $"📦 {objType.Name} ({count} items)"
        End If

        Dim identifier = GetObjectIdentifier(obj)
        Return $"🔗 {objType.Name}#{identifier}"
    End Function

    Private Function GetObjectIdentifier(obj As Object) As String
        If obj Is Nothing Then Return "null"

        Dim objType = obj.GetType()

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

        Return objType.Name
    End Function

    Private Function IsDefaultValue(value As Object, valueType As Type) As Boolean
        ' 🔹 Null, string gol sau colecții goale
        If value Is Nothing Then Return True
        If TypeOf value Is String AndAlso String.IsNullOrEmpty(DirectCast(value, String)) Then Return True
        If TypeOf value Is ICollection AndAlso DirectCast(value, ICollection).Count = 0 Then Return True
        If TypeOf value Is IDictionary AndAlso DirectCast(value, IDictionary).Count = 0 Then Return True

        ' 🔹 Boolean False (valoare implicită)
        If TypeOf value Is Boolean AndAlso DirectCast(value, Boolean) = False Then Return True

        ' 🔹 Alte valori implicite pentru tipuri value (0, 0.0, Date.MinValue, Enum=0 etc.)
        If valueType.IsValueType Then
            Try
                Dim defaultValue = Activator.CreateInstance(valueType)
                If value.Equals(defaultValue) Then Return True
            Catch
            End Try
        End If

        Return False
    End Function



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