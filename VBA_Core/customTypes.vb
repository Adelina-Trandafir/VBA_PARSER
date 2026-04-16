'PROJECT NAME: VBA_CORE
'FILE DESCRIPTION: Modul pentru tipuri de date utilizate în analiza VBA
'PATH: VBA_CORE/customTypes.vb

Imports System.Collections.Concurrent
Imports System.Runtime.CompilerServices
Imports System.Text

Public Module customObjects
    ''' <summary>
    ''' Dicționar personalizat care interceptează adăugarea, ștergerea și crearea de perechi.
    ''' </summary>
    <Serializable>
    Public Class CustomDict(Of TKey, TValue)
        Inherits Dictionary(Of TKey, TValue)

        ''' <summary>
        ''' Numele proprietății din obiectul părinte care deține acest dicționar.
        ''' </summary>
        Public Property OwnerName As String

        ''' <summary>
        ''' Constructor implicit – determină numele proprietății apelante (dacă e creat intern).
        ''' </summary>
        Public Sub New(<CallerMemberName> Optional propName As String = "")
            MyBase.New(If(GetType(TKey) Is GetType(String),
                      DirectCast(StringComparer.OrdinalIgnoreCase, IEqualityComparer(Of TKey)),
                      EqualityComparer(Of TKey).Default))
            OwnerName = propName
            LogInfo($"📘 CustomDict created for property '{OwnerName}' (type={GetType(TValue).Name})", 3)
        End Sub

        ''' <summary>
        ''' Setează explicit numele proprietății (folosit când dicționarul e atribuit manual din exterior).
        ''' </summary>
        Public Sub SetOwnerName(name As String)
            If String.IsNullOrEmpty(OwnerName) Then
                OwnerName = name
                LogInfo($"🧭 OwnerName assigned manually: '{OwnerName}' (type={GetType(TValue).Name})", 3)
            End If
        End Sub

        ''' <summary>
        ''' Adaugă o pereche cheie/valoare.  
        ''' Dacă cheia există deja, o modifică automat adăugând un sufix random (ex: #1, #2).
        ''' </summary>
        Public Shadows Sub Add(key As TKey, value As TValue)
            If key Is Nothing Then Exit Sub

            Dim finalKey As TKey = key
            Dim attempt As Integer = 1

            ' === 1️⃣ Evită coliziunea cu chei existente ===
            If MyBase.ContainsKey(finalKey) Then
                Do
                    Dim newKeyStr As String = $"{key}#{attempt}"
                    Try
                        ' Conversie sigură pentru tipul generic TKey
                        finalKey = DirectCast(Convert.ChangeType(newKeyStr, GetType(TKey)), TKey)
                    Catch
                        ' dacă TKey nu e convertibil din string, ieși elegant
                        LogError($"⚠️ Unable to auto-rename duplicate key '{key}' in dict '{OwnerName}'", New Exception("Type {GetType(TKey).Name} not string-convertible."))
                        Exit Do
                    End Try

                    attempt += 1
                Loop While MyBase.ContainsKey(finalKey)
            End If

            ' === 2️⃣ Adaugă efectiv ===
            LogInfo($"➕ Added key '{finalKey}' in dict '{OwnerName}' ({GetType(TValue).Name})", 3)
            MyBase.Add(finalKey, value)
        End Sub

        ''' <summary>
        ''' Elimină o cheie cu logare.
        ''' </summary>
        Public Shadows Function Remove(key As TKey) As Boolean
            LogInfo($"❌ Removed key '{key}' from dict '{OwnerName}' ({GetType(TValue).Name})", 3)
            Return MyBase.Remove(key)
        End Function

        ''' <summary>
        ''' Curăță toate elementele din dicționar cu logare.
        ''' </summary>
        Public Shadows Sub Clear()
            LogInfo($"🧹 Cleared dict '{OwnerName}' ({GetType(TValue).Name}) with {MyBase.Count} items", 3)
            MyBase.Clear()
        End Sub

        Public Overrides Function ToString() As String
            Return $"CustomDict '{OwnerName}' ({Count} items, type={GetType(TValue).Name})"
        End Function

        ''' <summary>
        ''' Permite accesul prin index numeric (0-based).
        ''' </summary>
        Default Public Overloads ReadOnly Property Item(index As Integer) As TValue
            Get
                If index < 0 OrElse index >= Me.Count Then
                    Throw New IndexOutOfRangeException($"Index {index} out of range for dict '{OwnerName}'.")
                End If
                Return Me.ElementAt(index).Value
            End Get
        End Property
    End Class

    ''' <summary>
    ''' Clasă de bază pentru liste personalizate, cu suport pentru OwnerName și logare la asignare/adăugare.
    ''' </summary>
    <Serializable>
    Public MustInherit Class TrackedList(Of T)
        Inherits List(Of T)

        Public Property OwnerName As String
        Private ReadOnly _instanceId As Guid = Guid.NewGuid()

        Public Sub New()
            LogInfo($"🧩 {GetType(T).Name} tracked list created (ID={_instanceId})", 3)
        End Sub

        Public Sub New(collection As IEnumerable(Of T))
            MyBase.New(collection)
            LogInfo($"🧩 {GetType(T).Name} tracked list cloned ({collection.Count}) (ID={_instanceId})", 3)
        End Sub

        Public Sub SetOwnerName(name As String)
            If String.IsNullOrEmpty(OwnerName) Then
                OwnerName = name
                LogInfo($"🧭 OwnerName assigned manually: '{OwnerName}' (ID={_instanceId})", 3)
            End If
        End Sub

        Public Overridable Sub OnAssignedToModule(moduleName As String)
            LogInfo($"📘 {GetType(T).Name} list '{OwnerName}' assigned to module '{moduleName}' (ID={_instanceId})", 3)
        End Sub
    End Class

    <Serializable>
    Public Class MethodLineList
        Inherits TrackedList(Of MethodLine)

        Public Sub New()
            LogInfo($"🧩 MethodLineList created", 3)
        End Sub

        Public Sub New(collection As IEnumerable(Of MethodLine))
            MyBase.New(collection)
            LogInfo($"🧩 MethodLineList created from existing collection ({collection.Count})", 3)
        End Sub

        Public Shadows Sub Add(item As MethodLine)
            If item Is Nothing Then Exit Sub
            If item.LocalLineNumber = 0 Then item.LocalLineNumber = Me.Count + 1
            LogInfo($"➕ Added line {item.LocalLineNumber} : {item.Content}", 3)
            MyBase.Add(item)
        End Sub

        Public Shadows Sub AddRange(collection As IEnumerable(Of MethodLine))
            If collection Is Nothing Then Exit Sub
            For Each ln In collection : Add(ln) : Next
            LogInfo(OwnerName & $" ➕ Added range of {collection.Count} lines", 3)
        End Sub

        Public Sub AddToRange(source As IList(Of MethodLine), startIndex As Integer, endIndex As Integer)
            If source Is Nothing OrElse startIndex < 0 OrElse endIndex >= source.Count OrElse startIndex > endIndex Then Exit Sub
            For i = startIndex To endIndex : Add(source(i)) : Next
            LogInfo(OwnerName & $" ➕ Added to range of lines from {startIndex} to {endIndex}", 3)
        End Sub
    End Class

    <Serializable>
    Public Class TokenList
        Inherits TrackedList(Of Token)

        Public Sub New()
            LogInfo($"🧩 TokenList created", 3)
        End Sub

        Public Sub New(collection As IEnumerable(Of Token))
            MyBase.New(collection)
            LogInfo($"🧩 TokenList created from existing collection ({collection.Count})", 3)
        End Sub

        Public Shadows Sub Add(item As Token)
            If item Is Nothing Then Exit Sub
            MyBase.Add(item)
            LogInfo($"➕ Added token '{item.TokenString}' : {item.TokenType}", 3)
        End Sub

        Public Shadows Sub AddRange(collection As IEnumerable(Of Token))
            If collection Is Nothing Then Exit Sub
            For Each t In collection : Add(t) : Next
            LogInfo(OwnerName & $" ➕ Added range of {collection.Count} tokens", 3)
        End Sub

        Public Sub AddToRange(source As IList(Of Token), startIndex As Integer, endIndex As Integer)
            If source Is Nothing OrElse startIndex < 0 OrElse endIndex >= source.Count OrElse startIndex > endIndex Then Exit Sub
            For i = startIndex To endIndex : Add(source(i)) : Next
            LogInfo(OwnerName & $" ➕ Added to range of tokens from {startIndex} to {endIndex}", 3)
        End Sub
    End Class

    <Serializable>
    Public Class ParamInfoList
        Inherits TrackedList(Of ParamInfo)

        Public Sub New()
            LogInfo($"🧩 ParamInfoList created", 3)
        End Sub

        Public Sub New(collection As IEnumerable(Of ParamInfo))
            MyBase.New(collection)
            LogInfo($"🧩 ParamInfoList created from existing collection ({collection.Count})", 3)
        End Sub

        Public Shadows Sub Add(item As ParamInfo)
            If item Is Nothing Then Exit Sub
            LogInfo($"➕ Added parameter {item}", 3)
            MyBase.Add(item)
        End Sub

        Public Shadows Sub AddRange(collection As IEnumerable(Of ParamInfo))
            If collection Is Nothing Then Exit Sub
            For Each p In collection : Add(p) : Next
            LogInfo(OwnerName & $" ➕ Added range of {collection.Count} parameters", 3)
        End Sub

        Public Sub AddToRange(source As IList(Of ParamInfo), startIndex As Integer, endIndex As Integer)
            If source Is Nothing OrElse startIndex < 0 OrElse endIndex >= source.Count OrElse startIndex > endIndex Then Exit Sub
            For i = startIndex To endIndex : Add(source(i)) : Next
            LogInfo(OwnerName & $" ➕ Added to range of parameters from {startIndex} to {endIndex}", 3)
        End Sub
    End Class

    <Serializable>
    Public Class EnumOrTypeList
        Inherits TrackedList(Of EnumOrType)

        Public Sub New()
            LogInfo($"🧩 EnumOrTypeList created", 3)
        End Sub

        Public Sub New(collection As IEnumerable(Of EnumOrType))
            MyBase.New(collection)
            LogInfo($"🧩 EnumOrTypeList created from existing collection ({collection.Count})", 3)
        End Sub

        Public Shadows Sub Add(item As EnumOrType)
            If item Is Nothing Then Exit Sub
            LogInfo($"➕ Added Enum/Type {item}", 3)
            MyBase.Add(item)
        End Sub

        Public Shadows Sub AddRange(collection As IEnumerable(Of EnumOrType))
            If collection Is Nothing Then Exit Sub
            For Each e In collection : Add(e) : Next
            LogInfo(OwnerName & $" ➕ Added range of {collection.Count} Enums/Types", 3)
        End Sub

        Public Sub AddToRange(source As IList(Of EnumOrType), startIndex As Integer, endIndex As Integer)
            If source Is Nothing OrElse startIndex < 0 OrElse endIndex >= source.Count OrElse startIndex > endIndex Then Exit Sub
            For i = startIndex To endIndex : Add(source(i)) : Next
            LogInfo(OwnerName & $" ➕ Added to range of Enums/Types from {startIndex} to {endIndex}", 3)
        End Sub
    End Class
End Module

Public Module customDesigner
    ''' <summary>
    ''' Reprezintă un element din structura raportului (ex. un control Begin/End din fișierul .txt al unui report Access).
    ''' </summary>
    <Serializable>
    Public Class FormReportDesigner
        Public Property Properties As FormReportDesignerPropertyList
        Public Property Children As FormReportDesignerChildrenList
        Public Property Parent As FormReportDesigner
        Public Property IsControl As Boolean
        Public Property IsSection As Boolean = True

        Public Sub New()
            Properties = New FormReportDesignerPropertyList(Me)
            Children = New FormReportDesignerChildrenList(Me)
        End Sub

        ''' <summary>
        ''' Permite acces direct la proprietățile controlului, ex: ctrl("Width"), ctrl("GUID").
        ''' Returnează lista valorilor (dacă există), altfel Nothing.
        ''' </summary>
        Default Public ReadOnly Property Item(propName As String) As String()
            Get
                If Properties Is Nothing Then Return Nothing
                Return Properties(propName)
            End Get
        End Property

        ''' <summary>
        ''' Numele nodului.  
        ''' Pentru controale vizuale (<see cref="IsControl"/> = True), returnează valoarea proprietății "Name".  
        ''' Pentru containere (secțiuni, grupuri etc.), returnează denumirea brută din fișier (ex. "BreakLevel", "PageHeader").
        ''' </summary>
        Public ReadOnly Property Name As String
            Get
                If IsControl Then
                    Dim n = Properties("Name")
                    If n IsNot Nothing AndAlso n.Length > 0 Then
                        Return n(0)
                    End If
                End If
                Return _rawName
            End Get
        End Property

        Private _rawName As String = ""
        ''' <summary>
        ''' Setează numele brut din fișierul de raport (nu cel al proprietății Access).
        ''' </summary>
        Public Property RawName As String
            Get
                Return _rawName
            End Get
            Set(value As String)
                _rawName = value
            End Set
        End Property

        ''' <summary>
        ''' Returnează valoarea unei proprietăți.
        ''' Dacă Full=True → concatenează toate liniile, eliminând ghilimelele de început/sfârșit.
        ''' Dacă ConcatWithString ≠ "", liniile sunt separate prin acel text.
        ''' </summary>
        Public ReadOnly Property GetProperty(propName As String,
                                             Optional Full As Boolean = False,
                                             Optional ConcatWithString As String = "") As String
            Get
                If Properties Is Nothing Then Return Nothing

                Dim vals = Properties(propName)
                If vals Is Nothing OrElse vals.Length = 0 Then Return Nothing

                If Not Full Then
                    Return vals(0)
                Else
                    Dim cleanedList As New List(Of String)
                    For Each v In vals
                        Dim cleaned = v.Trim()
                        If cleaned.StartsWith("""") Then cleaned = cleaned.Substring(1)
                        If cleaned.EndsWith("""") Then cleaned = cleaned.Substring(0, cleaned.Length - 1)
                        cleanedList.Add(cleaned)
                    Next

                    If String.IsNullOrEmpty(ConcatWithString) Then
                        Return String.Concat(cleanedList)
                    Else
                        Return String.Join(ConcatWithString, cleanedList)
                    End If
                End If
            End Get
        End Property

        Public Function GetTopParent() As FormReportDesigner
            Dim current As FormReportDesigner = Me
            While current?.Parent IsNot Nothing AndAlso current.Parent.Parent IsNot Nothing
                current = current.Parent
            End While
            Return current
        End Function

        ''' <summary>
        ''' Returnează o reprezentare JSON a elementului și a copiilor săi.
        ''' </summary>
        Public Function ToJson(Optional indentLevel As Integer = 0) As String
            Dim sb As New StringBuilder()
            Dim pad As New String(" "c, indentLevel * 2)

            sb.AppendLine(pad & "{")
            sb.AppendLine($"{pad}  ""Name"": ""{EscapeJson(Name)}"",")

            ' Proprietăți
            sb.AppendLine($"{pad}  ""Properties"": [")
            For i = 0 To Properties.Count - 1
                Dim prop = Properties.ElementAt(i)
                Dim comma = If(i < Properties.Count - 1, ",", "")
                sb.AppendLine($"{pad}    {{ ""Key"": ""{EscapeJson(prop.Key)}"", ""Values"": [{String.Join(", ", prop.Values.Select(Function(v) $"""{EscapeJson(v)}"""))}] }}{comma}")
            Next
            sb.AppendLine($"{pad}  ],")

            ' Copii
            sb.AppendLine($"{pad}  ""Children"": [")
            For i = 0 To Children.Count - 1
                Dim c = Children.ElementAt(i)
                Dim comma = If(i < Children.Count - 1, ",", "")
                sb.Append(c.ToJson(indentLevel + 2))
                If comma <> "" Then sb.AppendLine(",")
            Next
            sb.AppendLine()
            sb.AppendLine($"{pad}  ]")

            sb.Append(pad & "}")
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Scrie structura JSON direct în Debug.Print (frumos indentată).
        ''' </summary>
        Public Sub DumpJson()
            Debug.Print(ToJson())
        End Sub

        ''' <summary>
        ''' Returnează recursiv toate controalele vizuale (Label, TextBox, Line etc.) din acest nod și din descendenți.
        ''' </summary>
        Public Function GetAllControls() As List(Of FormReportDesigner)
            Dim result As New List(Of FormReportDesigner)

            ' Dacă elementul curent e control real → adaugă-l
            If IsControl AndAlso Not IsSection Then
                result.Add(Me)
            End If

            ' Caută recursiv în copii
            For Each child In Children
                result.AddRange(child.GetAllControls())
            Next

            Return result
        End Function

        ''' <summary>
        ''' Returnează toate controalele (FormReportDesigner) de pe nivelul curent,
        ''' fără a include recursiv copiii acestora.
        ''' </summary>
        Public Function GetControlsCurrentLevelOnly() As List(Of FormReportDesigner)
            Dim result As New List(Of FormReportDesigner)

            For Each child In Children
                If child.IsControl AndAlso Not IsSection Then
                    result.Add(child)
                End If
            Next

            Return result
        End Function

        Public Function GetAllSections() As List(Of FormReportDesigner)
            Dim result As New List(Of FormReportDesigner)

            ' Dacă elementul curent e control real → adaugă-l
            If IsSection AndAlso HasProperty("Name") Then
                result.Add(Me)
            End If

            ' Caută recursiv în copii
            For Each child In Children
                result.AddRange(child.GetAllSections())
            Next

            Return result
        End Function

        Public Function GetSectionsCurrentLevelOnly() As List(Of FormReportDesigner)
            Dim result As New List(Of FormReportDesigner)

            For Each child In Children
                If child.IsSection AndAlso child.HasProperty("Name") Then
                    result.Add(child)
                End If
            Next

            Return result
        End Function

        ''' <summary>
        ''' Verifică dacă acest element conține o proprietate cu numele dat (case-insensitive).
        ''' </summary>
        Public Function HasProperty(propName As String) As Boolean
            If String.IsNullOrEmpty(propName) OrElse Properties Is Nothing Then
                Return False
            End If
            Return Properties.Any(Function(p) String.Equals(p.Key, propName, StringComparison.OrdinalIgnoreCase))
        End Function

        ''' <summary>
        ''' Caută recursiv un control după valoarea proprietății "Name".
        ''' Returnează primul control găsit (sau Nothing dacă nu există).
        ''' </summary>
        ''' <param name="ctrlName">Valoarea proprietății "Name" a controlului căutat (ex: "Text107").</param>
        Public Function GetControl(ctrlName As String, Optional UseRawName As Boolean = False) As FormReportDesigner
            If IsControl Then
                ' 🔹 dacă elementul curent e control și are "Name" egal
                Dim nameProp As String
                If Not UseRawName Then
                    nameProp = GetProperty("Name")
                Else
                    nameProp = _rawName
                End If

                If nameProp <> "" AndAlso String.Equals(nameProp, ctrlName, StringComparison.OrdinalIgnoreCase) Then
                    Return Me
                End If
            End If

            ' 🔹 caută recursiv în copii
            For Each child In Children
                Dim found = child.GetControl(ctrlName, UseRawName)
                If found IsNot Nothing Then Return found
            Next

            Return Nothing
        End Function

        ''' <summary>
        ''' Returnează toate controalele (recursive) care conțin o anumită valoare
        ''' în oricare dintre proprietățile lor. Verifică toate valorile din fiecare String().
        ''' </summary>
        ''' <param name="value">Valoarea căutată (ex: "0x331a22eec1b22343a8c18608b7da4a88").</param>
        ''' <param name="exactMatch">Dacă True → match exact; dacă False → caută substring.</param>
        ''' <returns>Listă cu toate controalele (ReportDesigner) care conțin acea valoare.</returns>
        Public Function GetControlsByValue(value As String,
                                   Optional exactMatch As Boolean = True) As FormReportDesignerChildrenList

            Dim result As New FormReportDesignerChildrenList(Me)
            If String.IsNullOrEmpty(value) Then Return result

            Dim stack As New Stack(Of FormReportDesigner)
            stack.Push(Me)

            While stack.Count > 0
                Dim current = stack.Pop()
                If current Is Nothing Then Continue While

                If current.IsControl AndAlso current.Properties IsNot Nothing Then
                    For Each kvp In current.Properties
                        If kvp.Values IsNot Nothing Then
                            For Each v In kvp.Values
                                Dim match As Boolean = If(exactMatch,
                                                  v.Equals(value, StringComparison.OrdinalIgnoreCase),
                                                  v.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0)

                                If match Then
                                    result.Add(current)
                                    Exit For
                                End If
                            Next
                        End If
                        If result.Count > 0 AndAlso result.ElementAt(result.Count - 1) Is current Then Exit For
                    Next
                End If

                For Each child In current.Children
                    stack.Push(child)
                Next
            End While

            Return result
        End Function

        ''' <summary>
        ''' Returnează DOAR proprietățile controlului curent (fără copii) care
        ''' au o valoare egală / care conține 'value' în oricare dintre liniile String().
        ''' </summary>
        Public Function GetAllPropertiesByValue(value As String, Optional exactMatch As Boolean = True) As FormReportDesignerPropertyList
            Dim result As New FormReportDesignerPropertyList(Me)
            If String.IsNullOrEmpty(value) OrElse Properties Is Nothing OrElse Properties.Count = 0 Then
                Return result
            End If

            For Each kvp In Properties
                If kvp.Values Is Nothing Then Continue For

                Dim matchFound As Boolean = False
                For Each v As String In kvp.Values
                    If exactMatch Then
                        If v IsNot Nothing AndAlso v.Equals(value, StringComparison.OrdinalIgnoreCase) Then
                            matchFound = True
                            Exit For
                        End If
                    Else
                        If v IsNot Nothing AndAlso v.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0 Then
                            matchFound = True
                            Exit For
                        End If
                    End If
                Next

                If matchFound Then
                    ' asigură-te că treci String(), nu List(Of String)
                    Dim arr As String()

                    If TypeOf kvp.Values Is List(Of String) Then
                        arr = kvp.Values.ToArray()
                    Else
                        arr = CType(CType(kvp.Values, Object), String())
                    End If

                    result.Add(kvp.Key, arr)
                End If
            Next

            Return result
        End Function

        Private Function EscapeJson(value As String) As String
            If value Is Nothing Then Return ""
            Return value.Replace("\""", "\""").Replace("\", "\\")
        End Function
    End Class

    <Serializable>
    Public Class FormReportDesignerChildrenList
        Implements IEnumerable(Of FormReportDesigner)

        Private ReadOnly items As New List(Of FormReportDesigner)
        Private ReadOnly _owner As FormReportDesigner

        ''' <summary>
        ''' Elementul părinte al colecției (contextul arborelui).
        ''' </summary>
        Public ReadOnly Property Owner As FormReportDesigner
            Get
                Return _owner
            End Get
        End Property

        Public Sub New(owner As FormReportDesigner)
            Me._owner = owner
        End Sub

        ''' <summary>
        ''' Adaugă un copil nou și setează automat Parent.
        ''' </summary>
        Public Sub Add(child As FormReportDesigner)
            If child Is Nothing Then Exit Sub
            child.Parent = owner
            items.Add(child)
        End Sub

        ''' <summary>
        ''' Adaugă mai mulți copii (setează Parent pentru fiecare).
        ''' </summary>
        Public Sub AddRange(children As IEnumerable(Of FormReportDesigner))
            If children Is Nothing Then Exit Sub
            For Each ch In children
                If ch Is Nothing Then Continue For
                ch.Parent = owner
                items.Add(ch)
            Next
        End Sub

        ''' <summary>
        ''' Elimină un copil după referință.
        ''' </summary>
        Public Function Remove(child As FormReportDesigner) As Boolean
            Return items.Remove(child)
        End Function

        ''' <summary>
        ''' Elimină un copil după nume (primul match case-insensitive).
        ''' </summary>
        Public Function RemoveByName(name As String) As Boolean
            Dim found = items.FindIndex(Function(c) String.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase))
            If found >= 0 Then
                items.RemoveAt(found)
                Return True
            End If
            Return False
        End Function

        ''' <summary>
        ''' Golește colecția (nu modifică Parent-ul foștilor copii).
        ''' </summary>
        Public Sub Clear()
            items.Clear()
        End Sub

        ''' <summary>
        ''' Verifică existența unui copil cu numele dat (case-insensitive).
        ''' </summary>
        Public Function ContainsName(name As String) As Boolean
            Return items.Any(Function(c) String.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase))
        End Function

        ''' <summary>
        ''' Încearcă să obțină un copil după nume (case-insensitive).
        ''' </summary>
        Public Function TryGetByName(name As String, ByRef child As FormReportDesigner) As Boolean
            child = items.FirstOrDefault(Function(c) String.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase))
            Return child IsNot Nothing
        End Function

        ''' <summary>
        ''' Returnează lista numelor copiilor (în ordinea curentă).
        ''' </summary>
        Public ReadOnly Property Names As IReadOnlyList(Of String)
            Get
                Return items.Select(Function(c) c.Name).ToList()
            End Get
        End Property

        ''' <summary>
        ''' True dacă lista conține cel puțin un element.
        ''' </summary>
        Public ReadOnly Property Any As Boolean
            Get
                Return items.Count > 0
            End Get
        End Property

        ''' <summary>
        ''' Acces prin nume (case-insensitive). Aceasta este proprietatea implicită.
        ''' </summary>
        Default Public ReadOnly Property Item(name As String) As FormReportDesigner
            Get
                Return items.FirstOrDefault(Function(c) String.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase))
            End Get
        End Property

        ''' <summary>
        ''' Acces prin index numeric (0-based).
        ''' </summary>
        Public ReadOnly Property ItemAt(index As Integer) As FormReportDesigner
            Get
                Return items(index)
            End Get
        End Property

        Public ReadOnly Property Count As Integer
            Get
                Return items.Count
            End Get
        End Property

        ''' <summary>
        ''' Returnează elementul la index (alias pentru Item(index)).
        ''' </summary>
        Public Function ElementAt(index As Integer) As FormReportDesigner
            Return items(index)
        End Function

        ''' <summary>
        ''' Expune o copie a listei interne (read-only).
        ''' </summary>
        Public Function ToList() As List(Of FormReportDesigner)
            Return New List(Of FormReportDesigner)(items)
        End Function

        Public Function GetEnumerator() As IEnumerator(Of FormReportDesigner) Implements IEnumerable(Of FormReportDesigner).GetEnumerator
            Return items.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return GetEnumerator()
        End Function
    End Class

    <Serializable>
    Public Class FormReportDesignerProperty
        Public Property Key As String
        Public Property Values As List(Of String)
    End Class

    <Serializable>
    Public Class FormReportDesignerPropertyList
        Implements IEnumerable(Of FormReportDesignerProperty)

        Private ReadOnly items As New List(Of FormReportDesignerProperty)
        Private ReadOnly Owner As FormReportDesigner
        Public Sub New(owner As FormReportDesigner)
            Me.Owner = owner
        End Sub

        Public Sub Add(key As String, values As String())
            Dim cleaned = values.Select(Function(v) CleanQuotes(v)).ToList()
            Dim prop = New FormReportDesignerProperty With {.Key = key, .Values = cleaned}
            items.Add(prop)

            ' 🔹 dacă e proprietatea "Name", marchează părintele ca fiind control
            If key.Equals("Name", StringComparison.OrdinalIgnoreCase) Then
                If Owner IsNot Nothing Then Owner.IsControl = True
            ElseIf Key.Equals("Width", StringComparison.OrdinalIgnoreCase) Then
                ' 🔹 dacă proprietatea e "Height", atunci nu e secțiune
                If Owner IsNot Nothing Then Owner.IsSection = False
            End If
        End Sub

        Public Sub AddLineToProperty(line As String, Optional key As String = Nothing)
            If items.Count = 0 Then Return

            Dim target = If(String.IsNullOrEmpty(key),
                    items.Last(),
                    items.FirstOrDefault(Function(p) String.Equals(p.Key, key, StringComparison.OrdinalIgnoreCase)))

            If target Is Nothing Then Return
            target.Values.Add(CleanQuotes(line))

            ' 🔹 dacă proprietatea e "Name", marchează părintele ca fiind control
            If target.Key.Equals("Name", StringComparison.OrdinalIgnoreCase) Then
                If Owner IsNot Nothing Then Owner.IsControl = True
            ElseIf target.Key.Equals("Width", StringComparison.OrdinalIgnoreCase) Then
                ' 🔹 dacă proprietatea e "Height", atunci nu e secțiune
                If Owner IsNot Nothing Then Owner.IsSection = False
            End If
        End Sub

        Default Public ReadOnly Property Item(key As String) As String()
            Get
                Dim prop = items.FirstOrDefault(Function(p) String.Equals(p.Key, key, StringComparison.OrdinalIgnoreCase))
                Return prop?.Values.ToArray()
            End Get
        End Property

        Public ReadOnly Property Count As Integer
            Get
                Return items.Count
            End Get
        End Property

        Public Function ElementAt(index As Integer) As FormReportDesignerProperty
            Return items(index)
        End Function

        Public Function Contains(key As String) As Boolean
            Return items.Any(Function(p) String.Equals(p.Key, key, StringComparison.OrdinalIgnoreCase))
        End Function

        Public Function GetEnumerator() As IEnumerator(Of FormReportDesignerProperty) Implements IEnumerable(Of FormReportDesignerProperty).GetEnumerator
            Return items.GetEnumerator()
        End Function

        Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
            Return GetEnumerator()
        End Function

        Private Function CleanQuotes(value As String) As String
            If String.IsNullOrEmpty(value) Then Return value
            Dim r = value
            If r.StartsWith("""") Then r = r.Substring(1)
            If r.EndsWith("""") Then r = r.Substring(0, r.Length - 1)
            Return r
        End Function
    End Class

End Module

Public Module customTypes
    <Serializable>
    Public Class ModuleContainer
        ' =============================================================
        ' 🔹 METADATA
        ' =============================================================
        Public Property Name As String
        Public Property Type As String      ' "Form" / "Class" / "Module"
        Public Property CodeContent As String
        Public Property FilePath As String

        Private _WorkingLines As MethodLineList
        Public Property WorkingLines As MethodLineList
            Get
                Return _WorkingLines
            End Get
            Set(value As MethodLineList)
                If value IsNot Nothing Then value.SetOwnerName(NameOf(WorkingLines))
                _WorkingLines = value
            End Set
        End Property

        Private _Lines As MethodLineList
        Public Property Lines As MethodLineList
            Get
                Return _Lines
            End Get
            Set(value As MethodLineList)
                If value IsNot Nothing Then value.SetOwnerName(NameOf(Lines))
                _Lines = value
            End Set
        End Property

        Private _Constants As ParamInfoList
        Public Property Constants As ParamInfoList
            Get
                Return _Constants
            End Get
            Set(value As ParamInfoList)
                If value IsNot Nothing Then value.SetOwnerName(NameOf(Constants))
                _Constants = value
            End Set
        End Property

        Private _Variables As ParamInfoList
        Public Property Variables As ParamInfoList
            Get
                Return _Variables
            End Get
            Set(value As ParamInfoList)
                If value IsNot Nothing Then value.SetOwnerName(NameOf(Variables))
                _Variables = value
            End Set
        End Property

        Private _Types As EnumOrTypeList
        Public Property Types As EnumOrTypeList
            Get
                Return _Types
            End Get
            Set(value As EnumOrTypeList)
                If value IsNot Nothing Then value.SetOwnerName(NameOf(Types))
                _Types = value
            End Set
        End Property

        Private _Enums As EnumOrTypeList
        Public Property Enums As EnumOrTypeList
            Get
                Return _Enums
            End Get
            Set(value As EnumOrTypeList)
                If value IsNot Nothing Then value.SetOwnerName(NameOf(Enums))
                _Enums = value
            End Set
        End Property

        Private _Declares As CustomDict(Of String, MethodInfo)
        Public Property Declares As CustomDict(Of String, MethodInfo)
            Get
                Return _Declares
            End Get
            Set(value As CustomDict(Of String, MethodInfo))
                If value IsNot Nothing Then
                    value.SetOwnerName(NameOf(Declares))
                    LogInfo($"📗 Declares dict assigned to module '{Name}'", 3)
                End If
                _Declares = value
            End Set
        End Property

        Private _Events As CustomDict(Of String, MethodInfo)
        Public Property Events As CustomDict(Of String, MethodInfo)
            Get
                Return _Events
            End Get
            Set(value As CustomDict(Of String, MethodInfo))
                If value IsNot Nothing Then
                    value.SetOwnerName(NameOf(Events))
                    LogInfo($"📗 Events dict assigned to module '{Name}'", 3)
                End If
                _Events = value
            End Set
        End Property

        Private _Methods As CustomDict(Of String, MethodInfo)
        Public Property Methods As CustomDict(Of String, MethodInfo)
            Get
                Return _Methods
            End Get
            Set(value As CustomDict(Of String, MethodInfo))
                If value IsNot Nothing Then
                    value.SetOwnerName(NameOf(Methods))
                    LogInfo($"📗 Methods dict assigned to module '{Name}'", 3)
                End If
                _Methods = value
            End Set
        End Property

        Private _Properties As CustomDict(Of String, PropertyInfo)
        Public Property Properties As CustomDict(Of String, PropertyInfo)
            Get
                Return _Properties
            End Get
            Set(value As CustomDict(Of String, PropertyInfo))
                If value IsNot Nothing Then
                    value.SetOwnerName(NameOf(Properties))
                    LogInfo($"📗 Properties dict assigned to module '{Name}'", 3)
                End If
                _Properties = value
            End Set
        End Property

        Private _Functions As CustomDict(Of String, FunctionInfo)
        Public Property Functions As CustomDict(Of String, FunctionInfo)
            Get
                Return _Functions
            End Get
            Set(value As CustomDict(Of String, FunctionInfo))
                If value IsNot Nothing Then
                    value.SetOwnerName(NameOf(Functions))
                    LogInfo($"📗 Functions dict assigned to module '{Name}'", 3)
                End If
                _Functions = value
            End Set
        End Property

        Private _ModuleSymbols As CustomDict(Of String, SymbolEntry)
        Public Property ModuleSymbols As CustomDict(Of String, SymbolEntry)
            Get
                Return _ModuleSymbols
            End Get
            Set(value As CustomDict(Of String, SymbolEntry))
                If value IsNot Nothing Then
                    value.SetOwnerName(NameOf(ModuleSymbols))
                    LogInfo($"📗 ModuleSymbols dict assigned to module '{Name}'", 3)
                End If
                _ModuleSymbols = value
            End Set
        End Property

        ' === Graf de dependinte cross-module ===
        ''' <summary>Metodele din alte module apelate de acest modul. Format: "ModuleName:MethodName".</summary>
        Public Property CallsTo As New HashSet(Of String)
        ''' <summary>Metodele din alte module care apeleaza metode din acest modul. Format: "ModuleName:MethodName".</summary>
        Public Property CalledBy As New HashSet(Of String)

        Public _EventSets As New CustomDict(Of String, EventSet)
        Public Property EventSets As CustomDict(Of String, EventSet)
            Get
                Return _EventSets
            End Get
            Set(value As CustomDict(Of String, EventSet))
                If value IsNot Nothing Then
                    value.SetOwnerName(NameOf(ModuleSymbols))
                    LogInfo($"📗 Events dict assigned to module '{Name}'", 3)
                End If
                _EventSets = value
            End Set
        End Property

        Public Function GetMemberInfo(name As String) As (Kind As String, ReturnType As String)
            If String.IsNullOrEmpty(name) Then Return ("", "")

            Dim f As FunctionInfo = Nothing
            If Functions?.TryGetValue(name, f) Then
                Return ("Function", f.ReturnType)
            End If

            Dim p As PropertyInfo = Nothing
            If Properties?.TryGetValue(name, p) Then
                Return ($"Property_{p.PropertyType}", p.ReturnType)
            End If

            Dim v = Variables?.FirstOrDefault(Function(x) x.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
            If v IsNot Nothing Then
                Return ("Variable", v.Type)
            End If

            Dim c = Constants?.FirstOrDefault(Function(x) x.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
            If c IsNot Nothing Then
                Return ("Constant", c.Type)
            End If

            Return ("", "")
        End Function

        ''' <summary>
        ''' Intoarce doar liniile de cod care nu sunt tokenizate (WorkingLines).
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property GetWorkingLines() As MethodLineList
            Get
                Dim working As New MethodLineList
                If WorkingLines IsNot Nothing Then
                    For Each line In WorkingLines.Where(Function(l) Not l.IsTokenized)
                        working.Add(line)
                    Next
                End If
                Return working
            End Get
        End Property

        ''' <summary>
        ''' Returnează toate liniile de cod ne-tokenizate (WorkingLines) concatenate într-un singur String.
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property GetAllWorkingLines() As String
            Get
                Dim allLines As New StringBuilder
                If WorkingLines IsNot Nothing Then
                    For Each line In WorkingLines.Where(Function(l) Not l.IsTokenized)
                        allLines.AppendLine(line.Content)
                    Next
                End If

                Return allLines.ToString
            End Get
        End Property
    End Class

    <Serializable>
    Public Class FormReportContainer
        Inherits ModuleContainer

        Public Property RecordSource As String
        Public Property Tag As String
        Public Property Sections As New List(Of FormReportSection)
        Public Property Controls As New List(Of FormReportControl)
        Public Property DesignerContainer As FormReportDesigner
    End Class

    <Serializable>
    Public Class FormReportSection
        Public Property Name As String
        Public Property Type As String    ' ex: "Section", "FormHeader", "ReportFooter"
        Public Property Controls As New List(Of FormReportControl)
        Public Property Parent As New FormReportContainer
        Public Property DesignerControl As FormReportDesigner
    End Class

    <Serializable>
    Public Class EventSet
        Public Property EventName As String
        Public Property HandlerString As String
        Public Property RefObject As Object
        Public Property RefType As String

        ''' <summary>
        ''' Returnează concatenarea HandlerString și EventName separate prin underscore.
        ''' Dacă unul lipsește, îl omite elegant.
        ''' </summary>
        Public ReadOnly Property FullSignature As String
            Get
                If String.IsNullOrEmpty(HandlerString) Then Return EventName
                If String.IsNullOrEmpty(EventName) Then Return HandlerString
                Return $"{HandlerString}_{EventName}"
            End Get
        End Property

        ''' <summary>
        ''' Conversie implicită către String — returnează FullSignature.
        ''' </summary>
        Public Shared Widening Operator CType(e As EventSet) As String
            Return If(e Is Nothing, "", e.FullSignature)
        End Operator
    End Class

    <Serializable>
    Public Class FormReportControl
        Public Property Type As String = ""
        Public Property Name As String = ""
        Public Property ControlSource As String = ""
        Public Property RecordSource As String = ""
        Public Property RowSource As String = ""
        Public Property SourceObject As String = ""
        Public Property Caption As String = ""
        Public Property ParentSection As String = ""
        Public Property Section As Object
        Public Property Tag As String = ""
        Public Property SubLabel As FormReportControl
        Public Property DesignerControl As FormReportDesigner
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
        Public Property CodeContent As String
        Public Property IsDefault As Boolean
        Public Property IsEnumerable As Boolean
        Public Property IsAccessEventHandler As Boolean
        Public Property IsFunction As Boolean

        Public Property HandlerObject As String
        Public Property HandlerEvent As String
        Public Property HandlerObjectType As String
        Public Property ParsedLine As MethodLine

        Private _MethodLines As MethodLineList
        Public Property MethodLines As MethodLineList
            Get
                Return _MethodLines
            End Get
            Set(value As MethodLineList)
                If value IsNot Nothing Then value.SetOwnerName(NameOf(MethodLines))
                _MethodLines = value
                LogInfo($"📘 MethodLines assigned for method '{Name}'", 3)
            End Set
        End Property

        Private _WorkingLines As MethodLineList
        Public Property WorkingLines As MethodLineList
            Get
                Return _WorkingLines
            End Get
            Set(value As MethodLineList)
                If value IsNot Nothing Then value.SetOwnerName(NameOf(WorkingLines))
                _WorkingLines = value
            End Set
        End Property

        Private _Tokens As TokenList
        Public Property Tokens As TokenList
            Get
                Return _Tokens
            End Get
            Set(value As TokenList)
                If value IsNot Nothing Then value.SetOwnerName(NameOf(Tokens))
                _Tokens = value
                LogInfo($"📘 Tokens assigned for method '{Name}'", 3)
            End Set
        End Property

        Private _Parameters As ParamInfoList
        Public Property Parameters As ParamInfoList
            Get
                Return _Parameters
            End Get
            Set(value As ParamInfoList)
                If value IsNot Nothing Then value.SetOwnerName(NameOf(Parameters))
                _Parameters = value
                LogInfo($"📘 Parameters assigned for method '{Name}'", 3)
            End Set
        End Property

        Private _Constants As ParamInfoList
        Public Property Constants As ParamInfoList
            Get
                Return _Constants
            End Get
            Set(value As ParamInfoList)
                If value IsNot Nothing Then value.SetOwnerName(NameOf(Constants))
                _Constants = value
                LogInfo($"📘 Constants assigned for method '{Name}'", 3)
            End Set
        End Property

        Private _Variables As ParamInfoList
        Public Property Variables As ParamInfoList
            Get
                Return _Variables
            End Get
            Set(value As ParamInfoList)
                If value IsNot Nothing Then value.SetOwnerName(NameOf(Variables))
                _Variables = value
                LogInfo($"📘 Variables assigned for method '{Name}'", 3)
            End Set
        End Property

        Private _CodeLabels As TokenList
        Public Property CodeLabels As TokenList
            Get
                Return _CodeLabels
            End Get
            Set(value As TokenList)
                If value IsNot Nothing Then value.SetOwnerName(NameOf(CodeLabels))
                _CodeLabels = value
                LogInfo($"📘 CodeLabels assigned for method '{Name}'", 3)
            End Set
        End Property

        ''' <summary>
        ''' Intoarce doar liniile de cod care nu sunt tokenizate (WorkingLines).
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property GetWorkingLines() As MethodLineList
            Get
                Dim working As New MethodLineList
                If WorkingLines IsNot Nothing Then
                    For Each line In WorkingLines.Where(Function(l) Not l.IsTokenized)
                        working.Add(line)
                    Next
                End If
                Return working
            End Get
        End Property

        ''' <summary>
        ''' Returnează toate liniile de cod ne-tokenizate (WorkingLines) concatenate într-un singur String.
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property GetAllWorkingLines() As String
            Get
                Dim allLines As New StringBuilder
                If WorkingLines IsNot Nothing Then
                    For Each line In WorkingLines.Where(Function(l) Not l.IsTokenized)
                        allLines.AppendLine(line.Content)
                    Next
                End If

                Return allLines.ToString
            End Get
        End Property
    End Class

    <Serializable>
    Public Class FunctionInfo
        Inherits MethodInfo
        Public Property ReturnType As String
        Public Property ReturnObject As ParamInfo
    End Class

    <Serializable>
    Public Class PropertyInfo
        Inherits MethodInfo
        Public Property IsProperty As Boolean
        Public Property ReturnType As String
        Public Property ReturnObject As ParamInfo
        Public Property PropertyType As String      ' pentru proprietăți: Get / Let / Set
        Public Overrides Function ToString() As String
            Return $"{PropertyType} Property: {Name} As {ReturnType} (Lines: {StartLine}-{EndLine})"
        End Function
    End Class

    <Serializable>
    Public Class MethodLine
        Public Property LineNumber As Integer
        Public Property LocalLineNumber As Integer
        Public Property Content As String
        Public Property OriginalContent As String
        Public Property IsTokenized As Boolean = False

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
        Public Property IsMethodLine As Boolean
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

        Public Property IsLabelLine As Boolean
        Public Sub New()
            _Tokens = New TokenList()
        End Sub
        Private Property _tokens As TokenList
        Public Property Tokens As TokenList
            Get
                If _tokens Is Nothing Then _tokens = New TokenList
                Return _Tokens
            End Get
            Set(value As TokenList)
                If Not Object.ReferenceEquals(_Tokens, value) Then
                    ' doar loghează dacă cineva resetează lista complet
                    LogInfo($"Warning: Token list reassigned at line {LineNumber}", 3)
                End If
                _Tokens = value
            End Set
        End Property

        Private Property _preTokens As TokenList
        Public Property PreTokens As TokenList
            Get
                If _preTokens Is Nothing Then _preTokens = New TokenList
                Return _preTokens
            End Get
            Set(value As TokenList)
                If Not Object.ReferenceEquals(_preTokens, value) Then
                    ' doar loghează dacă cineva resetează lista complet
                    LogInfo($"Warning: Token list reassigned at line {LineNumber}", 3)
                End If
                _preTokens = value
            End Set
        End Property
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
        Public Property AssignedFrom As Token

        ' === Legaturi declaratie ↔ referinta ===
        ''' <summary>Lista tokenilor care referencuiesc aceasta declaratie (populata de BuildDeclarationLinks).</summary>
        Public Property References As New List(Of Token)
        ''' <summary>Tokenul de declaratie corespunzator acestei utilizari (populat de BuildDeclarationLinks).</summary>
        Public Property DeclarationToken As Token

        ' === Legaturi secventiale per linie ===
        ''' <summary>Urmatorul token din aceeasi linie sursa (in ordinea coloanelor).</summary>
        Public Property NextInLine As Token
        ''' <summary>Precedentul token din aceeasi linie sursa (in ordinea coloanelor).</summary>
        Public Property PrevInLine As Token

        Public Overrides Function ToString() As String
            Dim props As New List(Of String)

            ' 🔹 Denumirea tokenului
            If Not String.IsNullOrWhiteSpace(TokenString) Then
                props.Add($"Token='{TokenString}'")
            End If

            ' 🔹 Tipul tokenului (enum → string)
            'If TokenType <> TokenTypeEnum.unresolved Then
            props.Add($"Type={TokenType}")
            'End If

            ' 🔹 Tipul de date (dacă e declarat / detectat)
            If Not String.IsNullOrWhiteSpace(DataType) Then
                props.Add($"DataType={DataType}")
            End If

            ' 🔹 Context (numele metodei)
            If Not String.IsNullOrWhiteSpace(Context) Then
                props.Add($"Context={Context}")
            End If

            ' 🔹 Rezolvat / nerezolvat
            If IsResolved Then
                props.Add("Resolved=True")
            Else
                props.Add("Resolved=False")
            End If

            ' 🔹 Calificativ complet (ex: With / Parent chain)
            If Not String.IsNullOrWhiteSpace(QualifiedName) Then
                props.Add($"Qualified='{QualifiedName}'")
            End If

            ' 🔹 Poziționare globală și locală
            props.Add($"Global_Pos={LineNumber}:{ColumnNumber}")
            props.Add($"Method_Pos={MethodLineNumber}:{MethodColumnNumber}")

            ' 🔹 Opțional: scope / built-in
            If Not String.IsNullOrWhiteSpace(ResolvedScope) Then
                props.Add($"Scope={ResolvedScope}")
            End If
            If IsBuiltIn Then
                props.Add("BuiltIn=True")
            End If

            Return String.Join("; ", props)
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

    <Serializable>
    Public Class EnumOrType
        Public Property Name As String
        Public Property Type As String      ' "Enum" / "Type"
        Public Property IsPublic As Boolean
        Public Property IsGlobal As Boolean
        Public Property Members As New List(Of ParamInfo)
    End Class

    <Serializable>
    Public Class SymbolEntry
        Public Property Name As String
        Public Property TypeName As String
        Public Property DeclType As TokenTypeEnum
        Public Property Scope As String      ' "Public", "Private", "Friend"
        Public Property SourceObject As Object ' referință către obiectul real (ParamInfo, FunctionInfo, PropertyInfo)
        Public Property Category As String  ' "Local", "Global", "FormReportControl", etc.    
    End Class
End Module

Public Module modContainers
    ' Liste globale pentru clasificare token-uri
    Public GlobalModules As New ConcurrentDictionary(Of String, ModuleContainer)(StringComparer.OrdinalIgnoreCase)

    Public GlobalFormsReports As New ConcurrentDictionary(Of String, FormReportContainer)(StringComparer.OrdinalIgnoreCase)

    Public GlobalSources As New ConcurrentBag(Of (Name As String, Type As String, Code As String, FCI As Object))

    Public GlobalSymbols As New ConcurrentDictionary(Of String, SymbolEntry)(StringComparer.OrdinalIgnoreCase)

    Public GlobalDesigners As New ConcurrentDictionary(Of String, FormReportDesigner)(StringComparer.OrdinalIgnoreCase)
End Module

Public Module modClassFunctions
    Public Function FindSymbol(container As ModuleContainer, symbolName As String) As SymbolEntry
        If container Is Nothing OrElse container.ModuleSymbols Is Nothing OrElse String.IsNullOrWhiteSpace(symbolName) Then
            Return Nothing
        End If

        Dim sym As SymbolEntry = Nothing
        If container.ModuleSymbols.TryGetValue(symbolName, sym) Then
            Return sym
        End If

        Return Nothing
    End Function

    Public Function TryGetFormReportDesigner(key As String) As FormReportDesigner
        Dim obj As Object = Nothing
        Try
            If GlobalDesigners.TryGetValue(key, CType(obj, FormReportDesigner)) Then
                Return TryCast(obj, FormReportDesigner)
            End If
            Return Nothing

        Catch ex As Exception
            Debugger.Log(0, "Error", $"Error retrieving FormReportDesigner for key '{key}': {ex.Message}")
            Return Nothing
        End Try
    End Function
End Module