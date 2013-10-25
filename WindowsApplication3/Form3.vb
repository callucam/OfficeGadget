Public Class Form3
    Private Sub OK_Button_Click(sender As Object, e As EventArgs) Handles OK_Button.Click

        Dim i As Object
        Dim FirstName As String
        Dim LastName As String
        Dim Initials As String

        i = ComboBox1.SelectedIndex

        If i = -1 Then
            MsgBox("You must identify yourself to proceed.")
        Else
            FirstName = DataSet1.Tables(0).Rows(i).Item(0).ToString
            LastName = DataSet1.Tables(0).Rows(i).Item(1).ToString
            Initials = DataSet1.Tables(0).Rows(i).Item(2).ToString

            My.Settings.Initials = Initials
            My.Settings.FirstName = FirstName
            My.Settings.LastName = LastName

            Form1.Label24.Text = "Hello " & My.Settings.FirstName & "."

            My.Settings.InputDataFolder = Me.TextBox1.Text
            My.Settings.OutputDataFolder = Me.TextBox2.Text

            My.Settings.FavOne = FavOne.Text
            My.Settings.FavTwo = FavTwo.Text
            My.Settings.FavThree = FavThree.Text
            My.Settings.FavFour = FavFour.Text
            My.Settings.FavFive = FavFive.Text

            Form1.Fav1.Text = FavOne.Text
            Form1.Fav2.Text = FavTwo.Text
            Form1.Fav3.Text = FavThree.Text
            Form1.Fav4.Text = FavFour.Text
            Form1.Fav5.Text = FavFive.Text

            My.Settings.Save()

            MsgBox("Settings Saved")

            Application.Restart()

            Me.Close()

        End If

    End Sub

    Private Sub CloseButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        Me.Close()
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Label7.Text = "Version: " & System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString
        FavOne.Text = My.Settings.FavOne
        FavTwo.Text = My.Settings.FavTwo
        FavThree.Text = My.Settings.FavThree
        FavFour.Text = My.Settings.FavFour
        FavFive.Text = My.Settings.FavFive
        Me.TextBox1.Text = (My.Settings.InputDataFolder.ToString)
        Me.TextBox2.Text = (My.Settings.OutputDataFolder.ToString)
        Dim dt As DataTable
        Dim Additem2 As String
        Me.ComboBox1.Items.Clear()
        DataSet1.ReadXml(My.Settings.InputDataFolder & "\Employee List.xml")
        dt = DataSet1.Tables(0)
        For j = 0 To dt.Rows.Count - 1
            Additem2 = dt.Rows(j).Item(0).ToString & " " & dt.Rows(j).Item(1).ToString
            Me.ComboBox1.Items.Add(Additem2)
        Next

        If My.Settings.SetupRequired Then
        Else
            Me.ComboBox1.Text = My.Settings.FirstName & " " & My.Settings.LastName
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim MyFolderBrowser As New System.Windows.Forms.FolderBrowserDialog
        ' Description that displays above the dialog box control. 

        MyFolderBrowser.Description = "Select the Folder"
        ' Sets the root folder where the browsing starts from 
        MyFolderBrowser.RootFolder = Environment.SpecialFolder.MyComputer
        Dim dlgResult As DialogResult = MyFolderBrowser.ShowDialog()

        If dlgResult = Windows.Forms.DialogResult.OK Then
            TextBox1.Text = MyFolderBrowser.SelectedPath

        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim MyFolderBrowser As New System.Windows.Forms.FolderBrowserDialog
        ' Description that displays above the dialog box control. 

        MyFolderBrowser.Description = "Select the Folder"
        ' Sets the root folder where the browsing starts from 
        MyFolderBrowser.RootFolder = Environment.SpecialFolder.MyComputer
        Dim dlgResult As DialogResult = MyFolderBrowser.ShowDialog()

        If dlgResult = Windows.Forms.DialogResult.OK Then
            TextBox2.Text = MyFolderBrowser.SelectedPath

        End If

    End Sub

End Class