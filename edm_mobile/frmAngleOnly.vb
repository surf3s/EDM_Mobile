Public Class frmAngleOnly
    Inherits System.Windows.Forms.Form
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents foreshot As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Dim busy As Boolean
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        MyBase.Dispose(disposing)
    End Sub

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents backshot As System.Windows.Forms.TextBox
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.foreshot = New System.Windows.Forms.TextBox
        Me.backshot = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(224, 96)
        Me.Label1.Text = "Use this form to set the horizontal angle.  The foreshot angle will be uploaded. " & _
            " If you only know the backshot, the foreshot will be calculated. Enter angles as" & _
            " degrees minutes seconds."
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 107)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 24)
        Me.Label2.Text = "Foreshot angle :"
        '
        'foreshot
        '
        Me.foreshot.Location = New System.Drawing.Point(104, 107)
        Me.foreshot.Name = "foreshot"
        Me.foreshot.Size = New System.Drawing.Size(128, 21)
        Me.foreshot.TabIndex = 5
        '
        'backshot
        '
        Me.backshot.Location = New System.Drawing.Point(104, 154)
        Me.backshot.Name = "backshot"
        Me.backshot.Size = New System.Drawing.Size(128, 21)
        Me.backshot.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 154)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(96, 24)
        Me.Label3.Text = "Backshot angle :"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(136, 131)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(72, 16)
        Me.Label4.Text = "(ddd.mmss)"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(136, 178)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(72, 16)
        Me.Label5.Text = "(ddd.mmss)"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(8, 202)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(224, 35)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Set"
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.Add(Me.MenuItem2)
        '
        'MenuItem2
        '
        Me.MenuItem2.Text = "Close"
        '
        'frmAngleOnly
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(240, 268)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.backshot)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.foreshot)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmAngleOnly"
        Me.Text = "Set Horizontal Angle"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub foreshot_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles foreshot.KeyPress
        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) Or Asc(e.KeyChar) = 46 Or Asc(e.KeyChar) = 8 Then
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub foreshot_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles foreshot.TextChanged

        Dim a As Integer
        Dim degrees As String
        Dim therest As String
        Dim backshottxt As String

        If busy Then Exit Sub

        busy = True

        If foreshot.Text = "" Then
            backshot.Text = ""
            busy = False
            Exit Sub
        End If

        If Not IsNumeric(foreshot.Text) Then
            Exit Sub
        End If

        a = InStr(foreshot.Text, ".")
        If a <> 0 Then
            degrees = leftstr(foreshot.Text, a - 1)
            therest = Mid(foreshot.Text, a + 1)
        Else
            degrees = foreshot.Text
            therest = ""
        End If

        If CInt(Int(degrees)) >= 180 Then
            backshottxt = CStr(CInt(Int(degrees)) - 180) + "." + therest
        Else
            backshottxt = CStr(CInt(Int(degrees)) + 180) + "." + therest
        End If

        backshot.Text = backshottxt

        busy = False

    End Sub

    Private Sub backshot_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles backshot.TextChanged

        Dim a As Integer
        Dim degrees As String
        Dim therest As String
        Dim foreshottxt As String

        If busy Then Exit Sub

        busy = True

        If backshot.Text = "" Then
            foreshot.Text = ""
            busy = False
            Exit Sub
        End If

        a = InStr(backshot.Text, ".")
        If a <> 0 Then
            degrees = leftstr(backshot.Text, a - 1)
            therest = Mid(backshot.Text, a + 1)
        Else
            degrees = backshot.Text
            therest = ""
        End If

        If CInt(Int(degrees)) >= 180 Then
            foreshottxt = CStr(CInt(Int(degrees)) - 180) + "." + therest
        Else
            foreshottxt = CStr(CInt(Int(degrees)) + 180) + "." + therest
        End If

        foreshot.Text = foreshottxt

        busy = False

    End Sub

    Private Sub backshot_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles backshot.KeyPress
        If (Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 57) Or Asc(e.KeyChar) = 46 Or Asc(e.KeyChar) = 8 Then
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub MenuItem2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        Me.DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim angle As String
        Dim deg As Integer
        Dim min As Integer
        Dim sec As Single

        If Not IsNumeric(foreshot.Text) Then
            MsgBox("Enter a valid angle.", MsgBoxStyle.Information, "EDM-Mobile")
            Exit Sub
        End If

        Cursor.Current = Cursors.WaitCursor

        angle = foreshot.Text
        If InStr(angle, ".") = 0 Then
            angle = angle + "."
        End If

        deg = 0
        min = 0
        sec = 0
        Call sethortangle(angle, deg, min, sec)

        Cursor.Current = Cursors.Default

        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub
End Class
