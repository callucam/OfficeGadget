Public Class Form2
    Dim Y As Integer = Form1.DataGridView1.CurrentCellAddress.Y
    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Label6.Text = "Version: " & System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString
        Me.Text = "Time Record Number " & Y
        Label3.Text = "Details For Record Number " & Y
        StartTime.Value = Form1.DataSet1.Tables(0).Rows(Y).Item(0)
        EndTime.Value = Form1.DataSet1.Tables(0).Rows(Y).Item(1)
        Clientnumber.Text = Form1.DataSet1.Tables(0).Rows(Y).Item(3)
        ProjectNumber.Text = Form1.DataSet1.Tables(0).Rows(Y).Item(5)
        Task.Text = Form1.DataSet1.Tables(0).Rows(Y).Item(7)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        End
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form1.DataSet1.Tables(0).Rows(Y).Item(0) = StartTime.Value
        Form1.DataSet1.Tables(0).Rows(Y).Item(1) = EndTime.Value
        Form1.DataSet1.Tables(0).Rows(Y).Item(3) = Clientnumber.Text
        Form1.DataSet1.Tables(0).Rows(Y).Item(5) = ProjectNumber.Text
        Form1.DataSet1.Tables(0).Rows(Y).Item(7) = Task.Text
        MsgBox("Saved to table")
    End Sub
End Class