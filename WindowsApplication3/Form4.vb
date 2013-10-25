Public Class Form4



    'Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '    Dim Properties As Dictionary(Of Integer, KeyValuePair(Of String, String))
    '    Dim PropertyIndex As Integer
    '    ListBox2.Items.Clear()
    '    'For Each foundFile As String In My.Computer.FileSystem.GetFiles(My.Computer.FileSystem.SpecialDirectories.MyDocuments, FileIO.SearchOption.SearchAllSubDirectories, "*." & Extensionbox.Text)
    '    For Each foundFile As String In My.Computer.FileSystem.GetFiles(Folderbox.Text, FileIO.SearchOption.SearchAllSubDirectories, "*." & Extensionbox.Text)
    '        ListBox1.Items.Add(foundFile)
    '        Properties = GetFileProperties(foundFile)
    '        PropertyIndex = 0
    '        For Each FileProperty As KeyValuePair(Of Integer, KeyValuePair(Of String, String)) In Properties
    '            ListBox2.Items.Add(PropertyIndex & ": " & FileProperty.Value.Key & ": " & FileProperty.Value.Value)
    '            ListBox2.SetSelected(PropertyIndex, True)
    '            PropertyIndex = PropertyIndex + 1
    '        Next

    '    Next

    '    Dim S As String = ""
    '    For Each I As String In ListBox2.SelectedItems
    '        S &= I & vbCrLf
    '    Next

    '    'Clipboard.SetText(ListBox2.SelectedItems.ToString)
    '    If S <> "" Then
    '        Clipboard.SetText(S)
    '    End If
    '    '    ' Change the drive\path if necessary 
    '    '    Dim root As String = "C:\Program Files\Microsoft Visual Studio 9.0"

    '    '    'Take a snapshot of the folder contents 
    '    '    Dim dir As New System.IO.DirectoryInfo(root)

    '    '    Dim fileList = dir.GetFiles("*.*", System.IO.SearchOption.AllDirectories)

    '    '    ' This query will produce the full path for all .txt files 
    '    '    ' under the specified folder including subfolders. 
    '    '    ' It orders the list according to the file name. 
    '    '    Dim fileQuery = From file In fileList _
    '    '                    Where file.Extension = ".txt" _
    '    '                    Order By file.Name _
    '    '                    Select file

    '    '    For Each file In fileQuery
    '    '        Console.WriteLine(file.FullName)
    '    '    Next

    '    '    ' Create and execute a new query by using 
    '    '    ' the previous query as a starting point. 
    '    '    ' fileQuery is not executed again until the 
    '    '    ' call to Last 
    '    'Dim fileQuery2 = From file In fileQuery _
    '    '                 Order By file.CreationTime _
    '    '                 Select file.Name, file.CreationTime

    '    '    ' Execute the query 
    '    '    Dim newestFile = fileQuery2.Last

    '    '    Console.WriteLine("\r\nThe newest .txt file is {0}. Creation time: {1}", _
    '    '            newestFile.Name, newestFile.CreationTime)

    '    '    ' Keep the console window open in debug mode
    '    '    Console.WriteLine("Press any key to exit.")
    '    '    Console.ReadKey()










    'End Sub

    'Public Function GetFileProperties(ByVal FileName As String) As Dictionary(Of Integer, KeyValuePair(Of String, String))





    '    Dim Shell As New Shell
    '    Dim Folder As Folder = Shell.[NameSpace](Path.GetDirectoryName(FileName))
    '    Dim File As FolderItem = Folder.ParseName(Path.GetFileName(FileName))
    '    Dim Properties As New Dictionary(Of Integer, KeyValuePair(Of String, String))()
    '    Dim Index As Integer
    '    Dim Keys As Integer = Folder.GetDetailsOf(File, 0).Count
    '    'For Index = 0 To Keys - 1
    '    For Index = 0 To 500
    '        Dim CurrentKey As String = Folder.GetDetailsOf(Nothing, Index)
    '        Dim CurrentValue As String = Folder.GetDetailsOf(File, Index)
    '        'If CurrentValue <> "" Then
    '        Properties.Add(Index, New KeyValuePair(Of String, String)(CurrentKey, CurrentValue))
    '        'End If
    '    Next

    '    Return Properties





    'End Function




    ''Public ReadOnly Property ThisDrawing As AcadDocument
    ''    Get
    ''        Return Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.AcadDocument
    ''    End Get
    ''End Property


End Class