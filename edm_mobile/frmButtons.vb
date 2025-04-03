Public Class frmButtons
    Public buttonno As Integer

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click

        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub

    Private Sub frmButtons_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        buttonnolist.Text = "1"

        Dim a As Integer
        fieldlist.BeginUpdate()
        For a = 1 To vars
            Select Case UCase(varlist(a))
                Case "X", "Y", "Z", "HANGLE", "VANGLE", "DATUM", "SUFFIX"
                Case Else
                    fieldlist.Items.Add(varlist(a))
            End Select
        Next

        buttonno = 1
        Call readbutton(buttonno)

        fieldlist.EndUpdate()

    End Sub
    Private Sub readbutton(ByVal butno As Integer)

        Dim inidata(100, 2) As String
        Dim iniclass As String = "[BUTTON" & Trim(CStr(butno)) & "]"
        Dim status As Integer
        Dim a As Integer

        assignments.Text = ""
        inidata(1, 1) = "TITLE"
        For a = 1 To vars
            inidata(a + 1, 1) = varlist(a)
        Next a
        Call ReadIni(cfgfile, iniclass, inidata, status)
        If inidata(1, 2) <> "" Then
            Buttonpreview.Text = inidata(1, 2)
            buttontitletxt.Text = inidata(1, 2)
            For a = 1 To vars
                If inidata(a + 1, 2) <> "" Then
                    assignments.Text = assignments.Text + inidata(a + 1, 1) + "=" + inidata(a + 1, 2) + vbCrLf
                End If
            Next
        Else
            Buttonpreview.Text = ""
            buttontitletxt.Text = ""
        End If
        buttonnolist.Text = Trim(Str(butno))
        
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click

        If buttonno <> 1 Then
            buttonno = buttonno - 1
            Call readbutton(buttonno)
        End If

    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        If buttonno <> 6 Then
            buttonno = buttonno + 1
            Call readbutton(buttonno)
        End If

    End Sub

    Private Sub fieldlist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim t1 As String
        Dim t2 As String
        Dim a As Integer
        For a = 1 To vars
            If varlist(a) = fieldlist.SelectedItem.ToString Then
                If vtype(a) = 4 Then
                    defaultvaluelist.Visible = True
                    defaultvaluetxt.Visible = False
                    defaultvaluelist.BeginUpdate()
                    defaultvaluelist.Items.Clear()
                    t1 = vmenu(a)
                    Do Until t1 = ""
                        a = InStr(t1, ",")
                        If a <> 0 Then
                            t2 = leftstr(t1, a - 1)
                            If t2 <> "" Then
                                defaultvaluelist.Items.Add(t2)
                            End If
                            t1 = Mid(t1, a + 1)
                        ElseIf t1 <> "" Then
                            defaultvaluelist.Items.Add(t1)
                            Exit Do
                        End If
                    Loop
                    defaultvaluelist.EndUpdate()
                Else
                    defaultvaluelist.Visible = False
                    defaultvaluetxt.Visible = True
                End If
                Exit For
            End If
        Next

    End Sub

    Private Sub defaultvaluelist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If fieldlist.SelectedIndex = -1 Or defaultvaluelist.SelectedIndex = -1 Then
            MsgBox("Select a field and a value.", MsgBoxStyle.Information And MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        If buttontitletxt.Text = "" Then
            MsgBox("Provide a title before adding values.", MsgBoxStyle.Information And MsgBoxStyle.OkOnly)
            Exit Sub
        End If

        Dim iniclass As String
        Dim inidata(2, 2) As String
        Dim status As Integer
        Cursor.Current = Cursors.WaitCursor
        iniclass = "[Button" + Trim(Str(buttonno)) + "]"
        inidata(1, 1) = "Title"
        inidata(1, 2) = buttontitletxt.Text
        inidata(2, 1) = fieldlist.SelectedItem.ToString
        inidata(2, 2) = defaultvaluelist.SelectedItem.ToString
        Call WriteIni(cfgfile, iniclass, inidata, status)
        Call readbutton(buttonno)
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim a As Integer
        Dim iniclass As String
        Dim inidata(100, 2) As String
        Dim status As Integer
        Cursor.Current = Cursors.WaitCursor
        iniclass = "[Button" + Trim(Str(buttonno)) + "]"
        inidata(1, 1) = "Title"
        inidata(1, 2) = ""
        For a = 1 To vars
            inidata(a + 1, 1) = varlist(a)
            inidata(a + 1, 2) = ""
        Next
        Call WriteIni(cfgfile, iniclass, inidata, status)
        Call readbutton(buttonno)
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If defaultvaluetxt.Text = "" Then
            MsgBox("Provide a value first.", MsgBoxStyle.Information And MsgBoxStyle.OkOnly)
            Exit Sub
        End If
        If buttontitletxt.Text = "" Then
            MsgBox("Provide a title before adding values.", MsgBoxStyle.Information And MsgBoxStyle.OkOnly)
            Exit Sub
        End If

        Dim iniclass As String
        Dim inidata(2, 2) As String
        Dim status As Integer
        Cursor.Current = Cursors.WaitCursor
        iniclass = "[Button" + Trim(Str(buttonno)) + "]"
        inidata(1, 1) = "Title"
        inidata(1, 2) = buttontitletxt.Text
        inidata(2, 1) = fieldlist.SelectedItem.ToString
        inidata(2, 2) = defaultvaluetxt.Text
        Call WriteIni(cfgfile, iniclass, inidata, status)
        Call readbutton(buttonno)
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub buttontitletxt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Buttonpreview.Text = buttontitletxt.Text

    End Sub

    Private Sub fieldlist_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fieldlist.SelectedIndexChanged

    End Sub
End Class