Imports System.IO
Public Class frmBatteryTest
    Dim cancel_pressed As Boolean = False

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click

        Me.DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim whichprism As New frmWhichPrism
        Dim temp_donotwaitforprism As Boolean = donotwaitforprism
        Dim shot_count As Integer = 0

        Button1.Enabled = False
        donotwaitforprism = True
        shottype = -2               ' X-shot
        shot_count_txt.Text = ""
        shot_count_txt.Visible = True
        Dim f As String = Path.GetDirectoryName(cfgfile) + "\" + "shots.txt"

        Do Until whichprism.ShowDialog() = Windows.Forms.DialogResult.Cancel Or cancel_pressed
            shot_count += 1
            shot_count_txt.Text = Str(shot_count) + " shots have been recorded."

            Dim file3 As StreamWriter = File.CreateText(f)
            file3.WriteLine(shot_count_txt.Text)
            file3.Flush()
            file3.Close()
            file3.Dispose()
        Loop

        donotwaitforprism = temp_donotwaitforprism
        Button1.Enabled = True

    End Sub

End Class