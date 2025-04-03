<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class frmButtons
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Private mainMenu1 As System.Windows.Forms.MainMenu

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.mainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.Button2 = New System.Windows.Forms.Button
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.fieldlist = New System.Windows.Forms.ListBox
        Me.Buttonpreview = New System.Windows.Forms.Button
        Me.defaultvaluetxt = New System.Windows.Forms.TextBox
        Me.buttonnolist = New System.Windows.Forms.TextBox
        Me.buttontitletxt = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label6 = New System.Windows.Forms.Label
        Me.assignments = New System.Windows.Forms.TextBox
        Me.defaultvaluelist = New System.Windows.Forms.ListBox
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.SuspendLayout()
        '
        'mainMenu1
        '
        Me.mainMenu1.MenuItems.Add(Me.MenuItem1)
        Me.mainMenu1.MenuItems.Add(Me.MenuItem2)
        Me.mainMenu1.MenuItems.Add(Me.MenuItem3)
        '
        'MenuItem1
        '
        Me.MenuItem1.Text = "Previous"
        '
        'MenuItem2
        '
        Me.MenuItem2.Text = "Next Button"
        '
        'MenuItem3
        '
        Me.MenuItem3.Text = "Close"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.None
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(240, 265)
        Me.TabControl1.TabIndex = 37
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.Button2)
        Me.TabPage1.Controls.Add(Me.Label5)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.fieldlist)
        Me.TabPage1.Controls.Add(Me.Buttonpreview)
        Me.TabPage1.Controls.Add(Me.defaultvaluetxt)
        Me.TabPage1.Controls.Add(Me.buttonnolist)
        Me.TabPage1.Controls.Add(Me.buttontitletxt)
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.Button1)
        Me.TabPage1.Controls.Add(Me.Label6)
        Me.TabPage1.Controls.Add(Me.assignments)
        Me.TabPage1.Controls.Add(Me.defaultvaluelist)
        Me.TabPage1.Location = New System.Drawing.Point(0, 0)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(240, 242)
        Me.TabPage1.Text = "Buttons for Fields"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(171, 107)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(61, 27)
        Me.Button2.TabIndex = 47
        Me.Button2.Text = "Add"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(107, 62)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(55, 18)
        Me.Label5.Text = "Default :"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(6, 62)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(98, 18)
        Me.Label4.Text = "Field :"
        '
        'fieldlist
        '
        Me.fieldlist.Location = New System.Drawing.Point(6, 80)
        Me.fieldlist.Name = "fieldlist"
        Me.fieldlist.Size = New System.Drawing.Size(98, 86)
        Me.fieldlist.TabIndex = 43
        '
        'Buttonpreview
        '
        Me.Buttonpreview.Location = New System.Drawing.Point(168, 7)
        Me.Buttonpreview.Name = "Buttonpreview"
        Me.Buttonpreview.Size = New System.Drawing.Size(64, 64)
        Me.Buttonpreview.TabIndex = 40
        '
        'defaultvaluetxt
        '
        Me.defaultvaluetxt.Location = New System.Drawing.Point(108, 80)
        Me.defaultvaluetxt.Name = "defaultvaluetxt"
        Me.defaultvaluetxt.Size = New System.Drawing.Size(124, 21)
        Me.defaultvaluetxt.TabIndex = 39
        '
        'buttonnolist
        '
        Me.buttonnolist.Enabled = False
        Me.buttonnolist.Location = New System.Drawing.Point(95, 7)
        Me.buttonnolist.Name = "buttonnolist"
        Me.buttonnolist.Size = New System.Drawing.Size(67, 21)
        Me.buttonnolist.TabIndex = 36
        '
        'buttontitletxt
        '
        Me.buttontitletxt.Location = New System.Drawing.Point(95, 33)
        Me.buttontitletxt.Name = "buttontitletxt"
        Me.buttontitletxt.Size = New System.Drawing.Size(67, 21)
        Me.buttontitletxt.TabIndex = 35
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(7, 33)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(82, 18)
        Me.Label2.Text = "Button Title :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(7, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(82, 18)
        Me.Label1.Text = "Button No. :"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(11, 199)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(66, 40)
        Me.Button1.TabIndex = 31
        Me.Button1.Text = "Clear All"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(83, 178)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(98, 18)
        Me.Label6.Text = "Assignments :"
        '
        'assignments
        '
        Me.assignments.Location = New System.Drawing.Point(83, 199)
        Me.assignments.Multiline = True
        Me.assignments.Name = "assignments"
        Me.assignments.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.assignments.Size = New System.Drawing.Size(149, 40)
        Me.assignments.TabIndex = 30
        '
        'defaultvaluelist
        '
        Me.defaultvaluelist.Location = New System.Drawing.Point(108, 80)
        Me.defaultvaluelist.Name = "defaultvaluelist"
        Me.defaultvaluelist.Size = New System.Drawing.Size(124, 86)
        Me.defaultvaluelist.TabIndex = 44
        '
        'TabPage2
        '
        Me.TabPage2.Location = New System.Drawing.Point(0, 0)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(232, 239)
        Me.TabPage2.Text = "Buttons for Units"
        '
        'frmButtons
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(240, 268)
        Me.ControlBox = False
        Me.Controls.Add(Me.TabControl1)
        Me.Menu = Me.mainMenu1
        Me.Name = "frmButtons"
        Me.Text = "frmButtons"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents defaultvaluelist As System.Windows.Forms.ListBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents fieldlist As System.Windows.Forms.ListBox
    Friend WithEvents Buttonpreview As System.Windows.Forms.Button
    Friend WithEvents defaultvaluetxt As System.Windows.Forms.TextBox
    Friend WithEvents buttonnolist As System.Windows.Forms.TextBox
    Friend WithEvents buttontitletxt As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents assignments As System.Windows.Forms.TextBox
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
End Class
