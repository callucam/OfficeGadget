'before deployment 1. remove overwrite to setuprequired setting. 2. Set input/ output folders 3. set deployment folder


Imports System.Reflection
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Runtime.InteropServices
Imports System.Collections.Generic
Imports System.Xml
Imports System.Linq
Imports System.Xml.Linq
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms
Imports System.Data.OleDb

Public Class Form1
    Dim Projectno As String
    Dim Projectname As String
    Dim Clientnoitem As String
    Dim list As XElement
    Dim doc As XElement
    Dim records As XElement
    Dim projectnumbertemp As String
    Dim Additem2 As String
    Dim WhichWeek As String
    Dim WeekStart As Integer
    Dim WeekEnd As Integer
    Dim recordsAsDates As XElement
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If My.Computer.FileSystem.DirectoryExists(My.Settings.OutputDataFolder) And My.Computer.FileSystem.DirectoryExists(My.Settings.InputDataFolder) Then

        Me.Label28.Text = "Version: " & System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString
            Me.Label34.Text = My.Settings.OutputDataFolder & "\timerecord - " & My.Settings.Initials & ".xml"

        Dim i As Integer = 0
        Timer1.Enabled = False
        Me.Clientno.Text = ""
        Me.Clientno.Items.Clear()
        Projectno = ""
        Label38.Visible = False
        LineShape21.Visible = False

            Fav1.Text = My.Settings.FavOne
            Fav2.Text = My.Settings.FavTwo
            Fav3.Text = My.Settings.FavThree
            Fav4.Text = My.Settings.FavFour
            Fav5.Text = My.Settings.FavFive
 
        doc = XElement.Load(My.Settings.InputDataFolder & "\Project List.xml")
            records = XElement.Load(My.Settings.OutputDataFolder & "\timerecord - " & My.Settings.Initials & ".xml")
        GeneralReset()

        Me.ComboBox3.Items.Add("This Week")
        Me.ComboBox3.Items.Add("Last Week")
        Me.ComboBox3.Text = "This Week"
        WhichWeek = Me.ComboBox3.Text
        WeekStart = Int((CDbl(Now().ToOADate()) - 2) / 7) * 7 + 2
        WeekEnd = Int((CDbl(Now().ToOADate()) - 2) / 7) * 7 + 2 + 6

        If Math.Abs(records...<endclock>.Last.Value - records...<startclock>.Last.Value) * 1440 * 60 < 1 Then
            drawflag(records...<startclock>.Last.Value, records...<endclock>.Last.Value, records...<clientnumber>.Last.Value, records...<projectnumber>.Last.Value)
        End If

            '//// COMMENT THE NEXT LINE !!!!

            'My.Settings.SetupRequired = True

        If My.Settings.SetupRequired Then
            Dim frmChild As New Form3
            frmChild.Show()
            frmChild.BringToFront()
            frmChild.TopMost = True
            frmChild = Nothing
            My.Settings.SetupRequired = False
            My.Settings.Save()
        Else
            Label24.Text = "Hello " & My.Settings.FirstName & "."
        End If

        Else
            MsgBox("Input/ Output folders are not valid.  Please adjust application settings.")
        End If

    End Sub

    Private Sub cmdOK_Click(sender As Object, e As EventArgs) Handles cmdOK.Click
        Dim record1 As XElement
        Dim i As Integer
        Dim startclock As Double = CDbl(Now().ToOADate())
        Dim endclock As Double = CDbl(Now().ToOADate())
        Dim user As String = My.Settings.Initials
        Dim clientnumber As String = Clientno.Text
        Dim clientnm As String = Clientname.Text
        Dim projectnumber As String
        Dim projectnm As String
        Dim task As String
        Dim SRED As Boolean = chkSRED.Checked
        Dim TimeStart As Double
        Dim TimeEnd As Double
        If Clientno.Text = "" Or Clientno.Text = "Please Choose ... " Or ComboBox1.Text = "" Or ComboBox1.Text = "Please Choose ... " Then
            MsgBox("Please select a client number and a project number.")
        Else
            i = 0
            If Len(ComboBox1.Text) >= 3 Then
                projectnumber = Microsoft.VisualBasic.Left(ComboBox1.Text, 3)
            Else
                projectnumber = ""
            End If
            If Len(ComboBox1.Text) >= 4 Then
                projectnm = Microsoft.VisualBasic.Right(ComboBox1.Text, Len(ComboBox1.Text) - 4)
            Else
                projectnm = ""
            End If
            task = ComboBox2.Text
            TimeStart = CDbl(startclock)
            TimeEnd = CDbl(endclock)
            drawflag(TimeStart, TimeEnd, clientnumber, projectnumber)
            record1 =
    <record>
        <startclock><%= startclock %></startclock>
        <endclock><%= startclock %></endclock>
        <user><%= user %></user>
        <clientnumber><%= clientnumber %></clientnumber>
        <clientname><%= clientnm %></clientname>
        <projectnumber><%= projectnumber %></projectnumber>
        <projectname><%= projectnm %></projectname>
        <task><%= task %></task>
        <sred><%= SRED %></sred>
    </record>
            If Math.Abs(records...<endclock>.Last.Value - records...<startclock>.Last.Value) * 1440 * 60 < 1 Then
                records...<endclock>.Last.Value = endclock
            End If
            records.Add(record1)
            records.Save(My.Settings.OutputDataFolder & "\timerecord - " & My.Settings.Initials & ".xml")
            GeneralReset()
            Me.Refresh()
            End If
    End Sub

    Private Sub Clientno_TextChanged(sender As Object, e As EventArgs) Handles Clientno.TextChanged
        Dim i As Integer
        Dim Additem As String
        Dim nameString As String
        Me.ComboBox1.Items.Clear()
        Me.Clientname.Text = ""

        'For i = 0 To doc...<Row>.Count - 1
        '    nameString = doc...<clientnumber>.ToArray(i).Value
        '    If nameString = Clientno.Text Then
        '        Me.Clientname.Text = doc...<clientname>.ToArray(i).Value
        '        Additem = doc...<projectnumber>.ToArray(i).Value & " " & doc...<projectname>.ToArray(i).Value
        '        Me.ComboBox1.Items.Add(Additem)
        '    End If
        'Next

        Dim dt As DataTable
        DataSet3.Clear()
        DataSet3.ReadXml(My.Settings.InputDataFolder & "\Project List.xml")
        dt = DataSet3.Tables(0)
        For i = 0 To (dt.Rows.Count - 1)
            nameString = dt.Rows(i).Item(0).ToString
            If nameString = Clientno.Text Then
                Me.Clientname.Text = dt.Rows(i).Item(2).ToString
                Additem = dt.Rows(i).Item(1).ToString & " " & dt.Rows(i).Item(3).ToString
                Me.ComboBox1.Items.Add(Additem)
            End If
        Next

    End Sub

    Private Sub ComboBox1_Select(sender As Object, e As EventArgs) Handles ComboBox1.TextChanged
        Dim dt As DataTable
        Me.ComboBox2.Items.Clear()
        DataSet2.Clear()
        DataSet2.ReadXml(My.Settings.OutputDataFolder & "\timerecord - " & My.Settings.Initials & ".xml")
        dt = DataSet2.Tables(0)
        For i = (dt.Rows.Count - 1) To 0 Step -1
            If dt.Rows(i).Item(3).ToString = Clientno.Text And dt.Rows(i).Item(5).ToString = Microsoft.VisualBasic.Left(ComboBox1.Text, 3) Then
                Additem2 = dt.Rows(i).Item(7).ToString
                If Additem2 <> "" Then
                    Me.ComboBox2.Items.Add(Additem2)
                End If
            End If
        Next
    End Sub

    Private Sub groupBox1_Paint(ByVal sender As System.Object, ByVal ee As System.Windows.Forms.PaintEventArgs) Handles GroupBox1.Paint
        Dim upperleftx As Integer = 3
        Dim upperlefty As Integer = 3 + 12 + 1
        Dim lowerrightx As Integer = 710 + 3
        Dim lowerrighty As Integer = 140 + 3
        Dim SixAM As Integer = 150 + upperleftx
        Dim SixPM As Integer = 666 + upperleftx
        Dim TwelveHours As Integer = 666 - 150
        Dim M As Integer = 12 + upperlefty
        Dim T As Integer = 35 + upperlefty
        Dim W As Integer = 58 + upperlefty
        Dim Th As Integer = 81 + upperlefty
        Dim F As Integer = 104 + upperlefty
        Dim TimeStart As Double
        Dim TimeEnd As Double
        Dim LineStart As Integer
        Dim LineEnd As Integer
        Dim i As Integer
        Dim clientnumber As String
        Dim projectnumber As String
        Dim dow As Integer
        Dim dowposition As Integer
        Dim unique As Boolean = True
        Dim clientnumbers(0 To 99) As String
        Dim Arrayindex As Integer = 0
        Dim dt As DataTable
        DataSet2.Clear()
        DataSet2.ReadXml(My.Settings.OutputDataFolder & "\timerecord - " & My.Settings.Initials & ".xml")
        dt = DataSet2.Tables(0)
        For i = 0 To dt.Rows.Count - 1
            TimeStart = CDbl(dt.Rows(i).Item(0).ToString)
            If TimeStart > WeekStart And TimeStart < WeekEnd Then
                TimeEnd = CDbl(dt.Rows(i).Item(1).ToString)
                clientnumber = dt.Rows(i).Item(3).ToString
                projectnumber = dt.Rows(i).Item(5).ToString
                dow = Weekday(Date.FromOADate(TimeStart))
                If dow = vbSunday Then
                    dow = 8
                End If
                dowposition = (dow - 2) * 23 + 12 + upperlefty
                Using p As Pen = New Pen(Color.Black)
                    p.Width = 5
                    p.DashStyle = DashStyle.Solid
                    If clientnumber & "-" & projectnumber = Fav1.Text Then
                        p.Color = Color.Maroon
                    ElseIf clientnumber & "-" & projectnumber = Fav2.Text Then
                        p.Color = Color.SaddleBrown
                    ElseIf clientnumber & "-" & projectnumber = Fav3.Text Then
                        p.Color = Color.DarkOliveGreen
                    ElseIf clientnumber & "-" & projectnumber = Fav4.Text Then
                        p.Color = Color.Teal
                    ElseIf clientnumber & "-" & projectnumber = Fav5.Text Then
                        p.Color = Color.RoyalBlue
                    Else
                        p.Color = Color.Black
                    End If
                    p.EndCap = LineCap.ArrowAnchor
                    p.StartCap = LineCap.ArrowAnchor
                    LineStart = SixAM + TwelveHours * (-0.5 + 2 * (TimeStart - Int(TimeStart)))
                    LineEnd = SixAM + TwelveHours * (-0.5 + 2 * (TimeEnd - Int(TimeEnd)))
                    ee.Graphics.DrawLine(p, LineStart, dowposition, LineEnd, dowposition)
                End Using
            End If
        Next
    End Sub

    Private Sub Clockout_Click(sender As Object, e As EventArgs) Handles Clockout.Click
        Dim endclock As Double = CDbl(Now().ToOADate())
        records...<endclock>.Last.Value = endclock
        records.Save(My.Settings.OutputDataFolder & "\timerecord - " & My.Settings.Initials & ".xml")
        Timer1.Enabled = False
        Label38.Visible = False
        LineShape21.Visible = False

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If Label38.Visible = True Then 'checking whether it is visible
            Label38.Visible = False 'if visible then, make it invisible
            LineShape21.Visible = False
        ElseIf Label38.Visible = False Then 'checking whether it is invisible
            Label38.Visible = True 'if invisible then, make it visible
            LineShape21.Visible = True
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        WhichWeek = Me.ComboBox3.Text
        If WhichWeek = "This Week" Then
            WeekStart = Int((CDbl(Now().ToOADate()) - 2) / 7) * 7 + 2
            WeekEnd = Int((CDbl(Now().ToOADate()) - 2) / 7) * 7 + 2 + 6
            'Timer1.Enabled = True
            'Label38.Visible = True
            'LineShape21.Visible = True
        ElseIf WhichWeek = "Last Week" Then
            WeekStart = Int((CDbl(Now().ToOADate()) - 2) / 7) * 7 + 2 - 7
            WeekEnd = Int((CDbl(Now().ToOADate()) - 2) / 7) * 7 + 2 + 6 - 7
            Timer1.Enabled = False
            Label38.Visible = False
            LineShape21.Visible = False
        End If
        Me.Refresh()
    End Sub

    Private Sub CreateTimesheetButton_Click(sender As Object, e As EventArgs) Handles CreateTimesheetButton.Click
        Dim oXL As Excel.Application = Nothing
        Dim oWBs As Excel.Workbooks = Nothing
        Dim oWB2 As Excel.Workbook = Nothing
        Dim oSheet2 As Excel.Worksheet = Nothing
        Dim oCells As Excel.Range = Nothing
        Dim weekstartasdate As Date = CDate(Date.FromOADate(CDbl(WeekStart)))
        Dim SheetName As String = Year(weekstartasdate) & "-" & Month(weekstartasdate) & "-" & DateAndTime.Day(weekstartasdate)
        Dim timeDiff As Integer = 0
        Dim k As Integer = 0
        Dim unique As Integer
        Dim weekprojectno As String
        Dim weekclientname As String
        Dim weekprojectname As String
        Dim weektask As String
        Dim weektimediff As Double
        Dim SRED As String
        Dim currentrow As Integer = 12
        Dim i As Integer = 0
        Dim dt As DataTable
        oXL = New Excel.Application
        oXL.Visible = True
        oWBs = oXL.Workbooks
        oWB2 = oWBs.Open(My.Settings.InputDataFolder & "\XX 2000-01-01.xlsx")
        oSheet2 = oWB2.Worksheets(1)
        oSheet2.Range("b8").Offset(0, 0).Value = WeekStart
        oSheet2.Range("A7").Offset(0, 0).Value = My.Settings.FirstName & " " & My.Settings.LastName
        DataSet2.Clear()
        DataSet2.ReadXml(My.Settings.OutputDataFolder & "\timerecord - " & My.Settings.Initials & ".xml")
        dt = DataSet2.Tables(0)
        For d = 0 To 6
            For i = 0 To dt.Rows.Count - 1
                If Int(dt.Rows(i).Item(0).ToString) = Int(WeekStart) + d Then
                    weekprojectno = dt.Rows(i).Item(3).ToString & "-" & dt.Rows(i).Item(5).ToString
                    weekclientname = dt.Rows(i).Item(4).ToString
                    weekprojectname = dt.Rows(i).Item(6).ToString
                    weektask = dt.Rows(i).Item(7).ToString
                    weektimediff = Math.Round(4 * (dt.Rows(i).Item(1).ToString - dt.Rows(i).Item(0).ToString) * 24) / 4
                    unique = 1
                    SRED = dt.Rows(i).Item(8).ToString
                    If weektask <> "Clock-out" Then
                        For n = 0 To k
                            If oSheet2.Range("a11").Offset(n, 0).Value = weekprojectno Then
                                oSheet2.Range("e11").Offset(n, 0).Value = oSheet2.Range("e11").Offset(n, 0).Value & "; " & weektask
                                oSheet2.Range("f11").Offset(n, d).Value = oSheet2.Range("f11").Offset(n, d).Value + weektimediff
                                unique = 0
                            End If
                        Next
                        If unique = 1 Then
                            oSheet2.Rows(currentrow + k & ":" & currentrow + k).Insert()
                            oSheet2.Range("a11").Offset(k, 0).Value = weekprojectno
                            oSheet2.Range("b11").Offset(k, 0).Value = weekclientname
                            If SRED Then oSheet2.Range("c11").Offset(k, 0).Value = SRED Else oSheet2.Range("c11").Offset(k, 0).Value = ""
                            oSheet2.Range("d11").Offset(k, 0).Value = weekprojectname
                            oSheet2.Range("e11").Offset(k, 0).Value = weektask
                            oSheet2.Range("f11").Offset(k, d).Value = weektimediff
                            oSheet2.Range("m11").Offset(k, 0).Formula = "=SUM(F" & 11 + k & ":L" & 11 + k & ")"
                            k = k + 1
                        End If
                    End If
                End If
            Next
        Next
        oSheet2.PageSetup.LeftFooter = My.Settings.FirstName & " " & My.Settings.LastName
    End Sub

    Private Sub LoadButton_Click(sender As Object, e As EventArgs) Handles LoadButton.Click
        DataSet1.Clear()
        DataSet1.ReadXml(My.Settings.OutputDataFolder & "\timerecord - " & My.Settings.Initials & ".xml")
        For i = 0 To DataSet1.Tables(0).Rows.Count - 1
            DataSet1.Tables(0).Rows(i).Item(0) = CDate(Date.FromOADate(CDbl(DataSet1.Tables(0).Rows(i).Item(0).ToString)))
            DataSet1.Tables(0).Rows(i).Item(1) = CDate(Date.FromOADate(CDbl(DataSet1.Tables(0).Rows(i).Item(1).ToString)))
        Next
        DataGridView1.DataSource = Nothing
        DataGridView1.DataSource = DataSet1
        DataGridView1.DataMember = "record"
        DataGridView1.EditMode = DataGridViewEditMode.EditOnEnter
        DataGridView1.AllowUserToAddRows = True
        DataGridView1.AllowUserToDeleteRows = True
        DataGridView1.AllowUserToResizeColumns = False

    End Sub

    'Private Sub dataGridView1_CellClick(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
    '    Dim Y As Integer = DataGridView1.CurrentCellAddress.Y
    '    Dim TimeDuration As Double = ((records...<endclock>.ToArray(Y).Value) - (records...<startclock>.ToArray(Y).Value))
    '    Dim frmChild As New Form2
    '    frmChild.Show()
    '    frmChild = Nothing

    'End Sub

    'Private Sub CellValueChanged(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
    '    Dim fn As String = "Q:\Callum\Intranet\Gadget\" & My.Settings.Initials & "temp.xml"
    '    Dim records3 As XElement
    '    DataGridView1.Update()
    '    DataSet1 = DataGridView1.DataSource
    '    DataSet1.WriteXml(fn)
    '    records3 = XElement.Load(fn)
    '    For i = 0 To records3...<record>.Count - 1
    '        records3...<startclock>.ToArray(i).Value = CDbl(DateTime.Parse(records3...<startclock>.ToArray(i).Value).ToOADate())
    '        records3...<endclock>.ToArray(i).Value = CDbl(DateTime.Parse(records3...<endclock>.ToArray(i).Value).ToOADate())
    '    Next
    '    MsgBox("Change Recorded")
    '    records3.Save(My.Settings.OutputDataFolder & "\timerecord - " & My.Settings.Initials & ".xml")
    '    records = XElement.Load(My.Settings.OutputDataFolder & "\timerecord - " & My.Settings.Initials & ".xml")
    'End Sub

    'Private Sub LineShape12_Click(sender As Object, e As EventArgs) Handles LineShape12.Click

    '    ' Set the form as the parent of the ShapeContainer.
    '    canvas.Parent = Me.GroupBox1
    '    ' Set the ShapeContainer as the parent of the LineShape.
    '    lineA.Parent = canvas
    '    lineB.Parent = canvas

    '    lineA.StartPoint = New System.Drawing.Point(Me.Width / 2, 0)
    '    lineA.EndPoint = New System.Drawing.Point(Me.Width / 2, Me.Height)

    '    lineB.StartPoint = New System.Drawing.Point(Me.Width / 4, 0)
    '    lineB.EndPoint = New System.Drawing.Point(Me.Width / 4, Me.Height)

    '    MsgBox(canvas.Shapes.Count)
    'End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click
        Refresh()
    End Sub

    Private Sub FavRoutine(favprojectnumber As String, ClientnoText As String)

        Dim i As Integer
        Dim Additem As String
        Dim favclientnumber As String
        Me.Clientno.Items.Clear()
        Me.ComboBox1.Items.Clear()
        Me.ComboBox1.Text = favprojectnumber
        Me.Clientname.Text = ""
        Me.ComboBox2.Items.Clear()
        Me.ComboBox2.Text = ""

        Dim dt As DataTable
        DataSet3.Clear()
        DataSet3.ReadXml(My.Settings.InputDataFolder & "\Project List.xml")
        dt = DataSet3.Tables(0)
        For i = 0 To (dt.Rows.Count - 1)
            favclientnumber = dt.Rows(i).Item(0).ToString
            If favclientnumber = ClientnoText Then
                Me.Clientname.Text = dt.Rows(i).Item(2).ToString
                If favprojectnumber = dt.Rows(i).Item(1).ToString Then
                    ComboBox1.Items.Clear()
                    'projectnumbertemp = dt.Rows(i).Item(1).ToString
                    Additem = dt.Rows(i).Item(1).ToString & " " & dt.Rows(i).Item(3).ToString
                    'Me.ComboBox1.Items.Add(Additem)
                    Me.ComboBox1.Text = Additem
                    'ComboBox1.SelectedText = Additem
                End If

            End If
        Next

        If ComboBox1.Items.Count >= -1 Then
            ComboBox1.SelectedIndex = -1
        Else
        End If



        Me.ComboBox2.Focus()

    End Sub

    Private Sub Fav1_Click(sender As Object, e As EventArgs) Handles Fav1.Click

        Dim favprojectnumber As String = Microsoft.VisualBasic.Right(Fav1.Text, 3)
        Clientno.Text = (Microsoft.VisualBasic.Left(Fav1.Text, 3))
        FavRoutine(favprojectnumber, Clientno.Text)

    End Sub

    Private Sub Fav2_Click(sender As Object, e As EventArgs) Handles Fav2.Click

        Dim favprojectnumber As String = Microsoft.VisualBasic.Right(Fav2.Text, 3)
        Clientno.Text = (Microsoft.VisualBasic.Left(Fav2.Text, 3))
        FavRoutine(favprojectnumber, Clientno.Text)

    End Sub

    Private Sub Fav3_Click(sender As Object, e As EventArgs) Handles Fav3.Click

        Dim favprojectnumber As String = Microsoft.VisualBasic.Right(Fav3.Text, 3)
        Clientno.Text = (Microsoft.VisualBasic.Left(Fav3.Text, 3))
        FavRoutine(favprojectnumber, Clientno.Text)

    End Sub

    Private Sub Fav4_Click(sender As Object, e As EventArgs) Handles Fav4.Click

        Dim favprojectnumber As String = Microsoft.VisualBasic.Right(Fav4.Text, 3)
        Clientno.Text = (Microsoft.VisualBasic.Left(Fav4.Text, 3))
        FavRoutine(favprojectnumber, Clientno.Text)

    End Sub

    Private Sub Fav5_Click(sender As Object, e As EventArgs) Handles Fav5.Click

        Dim favprojectnumber As String = Microsoft.VisualBasic.Right(Fav5.Text, 3)
        Clientno.Text = (Microsoft.VisualBasic.Left(Fav5.Text, 3))
        FavRoutine(favprojectnumber, Clientno.Text)

    End Sub

    Private Sub PictureBox7_Click(sender As Object, e As EventArgs) Handles PictureBox7.Click
        If My.Computer.FileSystem.DirectoryExists("Q:\Callum\Intranet\Procedures") Then
            Process.Start("explorer.exe", "Q:\Callum\Intranet\Procedures")
        Else
            MsgBox("You do not appear to be connected to the the network Fraser.  Please connect and try again.")
        End If
    End Sub

    Private Sub PictureBox8_Click(sender As Object, e As EventArgs) Handles PictureBox8.Click

        If My.Computer.FileSystem.DirectoryExists("Q:\Callum\Intranet\Announcements") Then
            Process.Start("explorer.exe", "Q:\Callum\Intranet\Announcements")
        Else
            MsgBox("You do not appear to be connected to the the network Fraser.  Please connect and try again.")
        End If

    End Sub

    Private Sub PictureBox10_Click(sender As Object, e As EventArgs) Handles PictureBox10.Click

        If My.Computer.FileSystem.DirectoryExists("Q:\Callum\Intranet\Directory") Then
            Process.Start("explorer.exe", "Q:\Callum\Intranet\Directory")
        Else
            MsgBox("You do not appear to be connected to the the network Fraser.  Please connect and try again.")
        End If
    End Sub

    Private Sub PictureBox9_Click(sender As Object, e As EventArgs) Handles PictureBox9.Click

        MsgBox("Sorry, Search isn't yet ready for deployment.  Stay tuned.")

        'Dim frmChild As New Form4
        'frmChild.Show()
        'frmChild = Nothing

        'If My.Computer.FileSystem.DirectoryExists("Q:\Callum\Intranet\QA") Then
        '    Process.Start("explorer.exe", "Q:\Callum\Intranet\QA")
        'Else
        '    MsgBox("You do not appear to be connected to the the network Fraser.  Please connect and try again.")
        'End If

    End Sub

    Private Sub PictureBox12_Click(sender As Object, e As EventArgs) Handles PictureBox12.Click

        If My.Computer.FileSystem.DirectoryExists("Q:\Callum\Intranet\Calendar") Then
            Process.Start("explorer.exe", "Q:\Callum\Intranet\Calendar")
        Else
            MsgBox("You do not appear to be connected to the the network Fraser.  Please connect and try again.")
        End If

    End Sub

    Private Sub PictureBox11_Click(sender As Object, e As EventArgs) Handles PictureBox11.Click

        If My.Computer.FileSystem.DirectoryExists("Q:\Callum\Intranet\Resources") Then
            Process.Start("explorer.exe", "Q:\Callum\Intranet\Resources")
        Else
            MsgBox("You do not appear to be connected to the the network Fraser.  Please connect and try again.")
        End If

    End Sub

    Private Sub Other_Click_1(sender As Object, e As EventArgs) Handles Other.Click
        GeneralReset()
    End Sub

    Private Sub drawflag(TimeStart As Double, TimeEnd As Double, clientnumber As String, projectnumber As String)
        Dim newpoint As Point
        Dim xShift As Integer
        Dim yShift As Integer
        Dim dow As Integer
        Dim dowposition As Integer
        Dim LineEnd As Integer
        Dim upperleftx As Integer = 3
        Dim SixAM As Integer = 150 + upperleftx
        Dim TwelveHours As Integer = 666 - 150

        dow = Weekday(Date.FromOADate(TimeStart))
        If dow = vbSunday Then
            dow = 8
        End If
        dowposition = (dow - 2) * 23
        yShift = dowposition
        LineEnd = SixAM + TwelveHours * (-0.5 + 2 * (TimeEnd - Int(TimeEnd)))
        xShift = -153 + LineEnd
        newpoint.X = 151 + xShift
        newpoint.Y = 10 + yShift
        LineShape21.X1 = 151 + xShift
        LineShape21.X2 = 151 + xShift
        LineShape21.Y1 = 4 + yShift
        LineShape21.Y2 = 12 + yShift
        Label38.Location = newpoint
        Timer1.Interval = 500 'interval of the timer
        Timer1.Enabled = True 'starting the animation
        Label38.Text = clientnumber & "-" & projectnumber
        If clientnumber & "-" & projectnumber = Fav1.Text Then
            Label38.BackColor = Color.Maroon
            LineShape21.BorderColor = Color.Maroon
        ElseIf clientnumber & "-" & projectnumber = Fav2.Text Then
            Label38.BackColor = Color.SaddleBrown
            LineShape21.BorderColor = Color.SaddleBrown
        ElseIf clientnumber & "-" & projectnumber = Fav3.Text Then
            Label38.BackColor = Color.DarkOliveGreen
            LineShape21.BorderColor = Color.DarkOliveGreen
        ElseIf clientnumber & "-" & projectnumber = Fav4.Text Then
            Label38.BackColor = Color.Teal
            LineShape21.BorderColor = Color.Teal
        ElseIf clientnumber & "-" & projectnumber = Fav5.Text Then
            Label38.BackColor = Color.RoyalBlue
            LineShape21.BorderColor = Color.RoyalBlue
        Else
            Label38.BackColor = Color.Black
            LineShape21.BorderColor = Color.Black
        End If
    End Sub

    Private Sub GeneralReset()
        Me.Clientno.Items.Clear()
        Me.Clientname.Text = ""
        Me.ComboBox1.Text = "Please Choose ... "
        Me.Clientno.Text = "Please Choose ... "
        Me.ComboBox2.Items.Clear()
        Me.ComboBox2.Text = "Please Choose or Enter New ..."
        Me.chkSRED.Checked = False
        Me.Clientno.Focus()
        '    Dim b1 = From b In doc...<row>
        'Aggregate client In b...<clientname> _
        'Into ResultList = Any(True = True) _
        'Where ResultList = True
        '    For Each c In b1
        '        Clientno.Items.Add(c.b...<clientnumber>.Value)
        '    Next
        For i = 0 To doc...<Row>.Count - 1
            If doc...<isfirst>.ToArray(i).Value = "true" Then
                Clientno.Items.Add(doc...<clientnumber>.ToArray(i).Value)
            End If
        Next

    End Sub

    Private Sub SaveData_Click(sender As Object, e As EventArgs) Handles SaveData.Click
        'Dim fn As String = My.Settings.OutputDataFolder & "\timerecord - " & My.Settings.Initials & ".xml"
        'Dim records3 As XElement
        'DataSet1 = DataGridView1.DataSource
        'DataSet1.WriteXml(fn)
        'records3 = XElement.Load(fn)
        'For i = 0 To records3...<record>.Count - 1
        '    records3...<startclock>.ToArray(i).Value = CDbl(DateTime.Parse(records3...<startclock>.ToArray(i).Value).ToOADate())
        '    records3...<endclock>.ToArray(i).Value = CDbl(DateTime.Parse(records3...<endclock>.ToArray(i).Value).ToOADate())
        'Next
        'MsgBox("Change Recorded")
        'records3.Save(My.Settings.OutputDataFolder & "\timerecord - " & My.Settings.Initials & ".xml")
        'records = XElement.Load(My.Settings.OutputDataFolder & "\timerecord - " & My.Settings.Initials & ".xml")


        DataSet1.Tables(0).AcceptChanges()
        DataSet1.AcceptChanges()

        For i = 0 To DataSet1.Tables(0).Rows.Count - 1
            DataSet1.Tables(0).Rows(i).Item(0) = CDbl(DateTime.Parse(DataSet1.Tables(0).Rows(i).Item(0)).ToOADate())
            DataSet1.Tables(0).Rows(i).Item(1) = CDbl(DateTime.Parse(DataSet1.Tables(0).Rows(i).Item(1)).ToOADate())
        Next

        DataSet1.WriteXml(My.Settings.OutputDataFolder & "\timerecord - " & My.Settings.Initials & ".xml")
        records = XElement.Load(My.Settings.OutputDataFolder & "\timerecord - " & My.Settings.Initials & ".xml")
        MsgBox("Change Recorded")




    End Sub

    Private Sub Label25_Click(sender As Object, e As EventArgs) Handles Label25.Click
        Dim frmChild As New Form3
        frmChild.Show()
        frmChild = Nothing
    End Sub

    Private Sub Clientno_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Clientno.SelectedIndexChanged

    End Sub
End Class
