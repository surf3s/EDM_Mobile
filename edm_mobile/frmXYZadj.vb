Public Class frmXYZadj
    Inherits System.Windows.Forms.Form
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem

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
    Friend WithEvents movelabel As System.Windows.Forms.Label
    Friend WithEvents direction1 As System.Windows.Forms.RadioButton
    Friend WithEvents direction2 As System.Windows.Forms.RadioButton
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents units As System.Windows.Forms.ComboBox
    Friend WithEvents text1 As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button
        Me.movelabel = New System.Windows.Forms.Label
        Me.direction1 = New System.Windows.Forms.RadioButton
        Me.direction2 = New System.Windows.Forms.RadioButton
        Me.Label1 = New System.Windows.Forms.Label
        Me.units = New System.Windows.Forms.ComboBox
        Me.text1 = New System.Windows.Forms.TextBox
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(13, 170)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(214, 40)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "Move the point"
        '
        'movelabel
        '
        Me.movelabel.Location = New System.Drawing.Point(83, 45)
        Me.movelabel.Name = "movelabel"
        Me.movelabel.Size = New System.Drawing.Size(96, 19)
        Me.movelabel.Text = "Move the point"
        Me.movelabel.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'direction1
        '
        Me.direction1.Location = New System.Drawing.Point(64, 76)
        Me.direction1.Name = "direction1"
        Me.direction1.Size = New System.Drawing.Size(72, 26)
        Me.direction1.TabIndex = 4
        Me.direction1.Text = "Up"
        '
        'direction2
        '
        Me.direction2.Location = New System.Drawing.Point(136, 76)
        Me.direction2.Name = "direction2"
        Me.direction2.Size = New System.Drawing.Size(72, 26)
        Me.direction2.TabIndex = 3
        Me.direction2.Text = "Down"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(32, 120)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 24)
        Me.Label1.Text = "by"
        '
        'units
        '
        Me.units.Items.Add("M")
        Me.units.Items.Add("CM")
        Me.units.Items.Add("MM")
        Me.units.Location = New System.Drawing.Point(136, 120)
        Me.units.Name = "units"
        Me.units.Size = New System.Drawing.Size(72, 22)
        Me.units.TabIndex = 1
        '
        'text1
        '
        Me.text1.Location = New System.Drawing.Point(64, 120)
        Me.text1.Name = "text1"
        Me.text1.Size = New System.Drawing.Size(64, 21)
        Me.text1.TabIndex = 0
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.Add(Me.MenuItem1)
        '
        'MenuItem1
        '
        Me.MenuItem1.Text = "Cancel"
        '
        'frmXYZadj
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(240, 268)
        Me.ControlBox = False
        Me.Controls.Add(Me.text1)
        Me.Controls.Add(Me.units)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.direction2)
        Me.Controls.Add(Me.direction1)
        Me.Controls.Add(Me.movelabel)
        Me.Controls.Add(Me.Button1)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmXYZadj"
        Me.Text = "Move point"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim adj As Double

        If Not direction1.Checked And Not direction2.Checked Then
            MsgBox("Select a direction to the move the point.", MsgBoxStyle.Information And MsgBoxStyle.OkOnly, "EDM-Mobile")
            Exit Sub
        End If

        If text1.Text = "" Then
            MsgBox("Enter an adjustment amount.", MsgBoxStyle.Information And MsgBoxStyle.OkOnly, "EDM-Mobile")
            Exit Sub
        End If

        If Not IsNumeric(text1.Text) Then
            MsgBox("Enter an adjustment amount.", MsgBoxStyle.Information And MsgBoxStyle.OkOnly, "EDM-Mobile")
            Exit Sub
        End If

        Select Case LCase(units.Text)
            Case "m"
                adj = CDbl(text1.Text)
            Case "cm"
                adj = CDbl(text1.Text) / 100
            Case "mm"
                adj = CDbl(text1.Text) / 1000
            Case Else
                MsgBox("Select the units.", MsgBoxStyle.Information And MsgBoxStyle.OkOnly, "EDM-Mobile")
                Exit Sub
        End Select

        If direction1.Checked = True Then adj = adj * -1

        If MsgBox("Adjust point by " + CStr(adj) + " meters?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "EDM-Mobile") = MsgBoxResult.No Then Exit Sub

        measurementadjustment = adj
        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click

        Me.DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub

    Private Sub frmXYZadj_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        units.Text = "CM"

    End Sub
End Class
