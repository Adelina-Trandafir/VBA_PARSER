Imports System.Collections.Concurrent
Imports System.Reflection
Imports System.Threading
Imports System.Threading.Tasks
Imports VBA_CORE
Imports VBA_CORE.customTypes
Imports VBA_TOKENIZER  ' Adăugat conform cererii tale

Class MainWindow

    ' === MODEL NON-UI PENTRU ARBORE ===
    Public Class TreeNodeData
        Public Property Header As String
        Public Property Children As New List(Of TreeNodeData)
    End Class

    Private Class GlobalItem
        Public Property Name As String
        Public Property Data As Object
        Public Overrides Function ToString() As String
            Return Name
        End Function
    End Class

    Private _cancellationToken As CancellationTokenSource
    Private _visited As New ConcurrentDictionary(Of Object, Byte)

    Public Sub New()
        InitializeComponent()
        Program.RunFullAnalysis()
        LoadGlobalCollections()
    End Sub

    Private Sub LoadGlobalCollections()
        Dim items As New List(Of GlobalItem) From {
            New GlobalItem With {.Name = "GlobalModules", .Data = modContainers.GlobalModules},
            New GlobalItem With {.Name = "GlobalFormsReports", .Data = modContainers.GlobalFormsReports}
        }
        cboGlobals.ItemsSource = items
        cboGlobals.SelectedIndex = 0
    End Sub

    Private Async Sub cboGlobals_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        Await RefreshDataAsync()
    End Sub

    Private Async Sub btnRefresh_Click(sender As Object, e As RoutedEventArgs)
        Await RefreshDataAsync()
    End Sub

    Private Async Function RefreshDataAsync() As Task
        Dim item = TryCast(cboGlobals.SelectedItem, GlobalItem)
        If item Is Nothing Then Return

        ' Reset UI
        tvData.Items.Clear()
        progressBar.Value = 0
        btnCancel.IsEnabled = True

        _cancellationToken?.Cancel()
        _cancellationToken = New CancellationTokenSource()
        _visited.Clear()

        Dim data = item.Data
        Dim rootHeader = $"{item.Name} ({If(data?.GetType().Name, "Nothing")})"

        Dim rootNode As TreeViewItem = Nothing
        Dispatcher.Invoke(Sub()
                              rootNode = New TreeViewItem With {.Header = rootHeader, .IsExpanded = True}
                              tvData.Items.Add(rootNode)
                          End Sub)

        If data Is Nothing Then
            Dispatcher.Invoke(Sub() rootNode.Items.Add(New TreeViewItem With {.Header = "Data = Nothing"}))
            btnCancel.IsEnabled = False
            Return
        End If

        Dim dict = TryCast(data, IDictionary)
        If dict Is Nothing Then
            Dispatcher.Invoke(Sub() rootNode.Items.Add(New TreeViewItem With {.Header = "Not a dictionary"}))
            btnCancel.IsEnabled = False
            Return
        End If

        Dim total = dict.Count
        If total = 0 Then
            Dispatcher.Invoke(Sub() rootNode.Items.Add(New TreeViewItem With {.Header = "Empty collection"}))
            progressBar.Value = 100
            btnCancel.IsEnabled = False
            Return
        End If

        Dim processed As Integer = 0

        ' === PARALEL: Construiește DATE non-UI ===
        Dim tasks = dict.Keys.Cast(Of Object).Select(Function(key)
                                                         Return Task.Run(Function()
                                                                             If _cancellationToken.Token.IsCancellationRequested Then Return Nothing
                                                                             Try
                                                                                 Dim value = dict(key)
                                                                                 Dim keyStr = If(key?.ToString(), "<null>")
                                                                                 Dim dataNode = BuildTreeData_Parallel($"[{keyStr}]", value)
                                                                                 Interlocked.Increment(processed)
                                                                                 Return New KeyValuePair(Of String, TreeNodeData)(keyStr, dataNode)
                                                                             Catch
                                                                                 Return New KeyValuePair(Of String, TreeNodeData)(key?.ToString(), New TreeNodeData With {.Header = "[ERROR]"})
                                                                             End Try
                                                                         End Function, _cancellationToken.Token)
                                                     End Function).ToArray()

        Dim results = Await Task.WhenAll(tasks)

        ' === ADAUGĂ în UI thread: Convertește la TreeViewItem ===
        Dispatcher.Invoke(Sub()
                              For Each kvp In results.Where(Function(x) x.Value IsNot Nothing)
                                  Dim uiNode = ConvertToTreeViewItem(kvp.Value)
                                  rootNode.Items.Add(uiNode)
                              Next
                              progressBar.Value = 100
                              btnCancel.IsEnabled = False
                          End Sub)

        ' Progres live (buclă separată)
        Dim progressTask = Task.Run(Sub()
                                        While processed < total AndAlso Not _cancellationToken.Token.IsCancellationRequested
                                            Dim percent = If(total > 0, (processed / total) * 100, 100)
                                            Dispatcher.Invoke(Sub() progressBar.Value = percent)
                                            Thread.Sleep(100)  ' Update la 0.1 sec
                                        End While
                                    End Sub)
    End Function

    Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs)
        _cancellationToken?.Cancel()
        btnCancel.IsEnabled = False
    End Sub

    ' === CONSTRUIEȘTE DATE NON-UI (paralel-safe) ===
    Private Function BuildTreeData_Parallel(label As String, value As Object) As TreeNodeData
        If value Is Nothing Then
            Return New TreeNodeData With {.Header = $"{label} = Nothing"}
        End If

        Dim typeName = value.GetType().Name
        If typeName.Contains("String") Then typeName = "String"

        If Not _visited.TryAdd(value, 0) Then
            Return New TreeNodeData With {.Header = $"{label} [Circular Reference]"}
        End If

        Dim node As New TreeNodeData With {.Header = $"{label} : {typeName}"}

        Try
            If TypeOf value Is IDictionary Then
                Dim dict = DirectCast(value, IDictionary)
                For Each k In dict.Keys
                    If _cancellationToken?.Token.IsCancellationRequested Then Exit For
                    Dim v = dict(k)
                    node.Children.Add(BuildTreeData_Parallel($"[{k}]", v))
                Next

            ElseIf TypeOf value Is IEnumerable AndAlso Not TypeOf value Is String Then
                Dim list = DirectCast(value, IEnumerable)
                Dim i = 0
                For Each item In list
                    If _cancellationToken?.Token.IsCancellationRequested Then Exit For
                    node.Children.Add(BuildTreeData_Parallel($"[{i}]", item))
                    i += 1
                Next

            ElseIf value.GetType().IsClass Then
                Dim props = value.GetType().GetProperties(BindingFlags.Public Or BindingFlags.Instance)
                For Each p In props
                    If _cancellationToken?.Token.IsCancellationRequested Then Exit For
                    Dim propVal As Object = Nothing
                    Try
                        propVal = p.GetValue(value)
                    Catch
                        propVal = "<error>"
                    End Try
                    node.Children.Add(BuildTreeData_Parallel(p.Name, propVal))
                Next
            Else
                Dim display = If(TypeOf value Is String, $"""{value}""", value.ToString())
                node.Header = $"{label} = {display}"
            End If
        Finally
            _visited.TryRemove(value, Nothing)
        End Try

        Return node
    End Function

    ' === CONVERTEȘTE LA UI (în UI thread) ===
    Private Function ConvertToTreeViewItem(data As TreeNodeData) As TreeViewItem
        Dim item As New TreeViewItem With {.Header = data.Header}
        For Each childData In data.Children
            item.Items.Add(ConvertToTreeViewItem(childData))
        Next
        Return item
    End Function

End Class