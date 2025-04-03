Public Class frmBoot
    Inherits System.Windows.Forms.Form
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents xvalue As System.Windows.Forms.Label
    Friend WithEvents yvalue As System.Windows.Forms.Label
    Friend WithEvents zvalue As System.Windows.Forms.Label
    Friend WithEvents cfgstatus As System.Windows.Forms.Label
    Friend WithEvents dbstatus As System.Windows.Forms.Label
    Friend WithEvents logstatus As System.Windows.Forms.Label
    Friend WithEvents recstatus As System.Windows.Forms.Label

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
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.xvalue = New System.Windows.Forms.Label
        Me.yvalue = New System.Windows.Forms.Label
        Me.zvalue = New System.Windows.Forms.Label
        Me.cfgstatus = New System.Windows.Forms.Label
        Me.dbstatus = New System.Windows.Forms.Label
        Me.recstatus = New System.Windows.Forms.Label
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.logstatus = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(176, 16)
        Me.Label1.Text = "Current Station Coordinates:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(32, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(24, 16)
        Me.Label2.Text = "X :"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(32, 48)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(24, 16)
        Me.Label3.Text = "Y :"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(32, 64)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(24, 16)
        Me.Label4.Text = "Z :"
        '
        'xvalue
        '
        Me.xvalue.Location = New System.Drawing.Point(64, 32)
        Me.xvalue.Name = "xvalue"
        Me.xvalue.Size = New System.Drawing.Size(112, 16)
        Me.xvalue.Text = "12345678.123"
        '
        'yvalue
        '
        Me.yvalue.Location = New System.Drawing.Point(64, 48)
        Me.yvalue.Name = "yvalue"
        Me.yvalue.Size = New System.Drawing.Size(112, 16)
        Me.yvalue.Text = "12345678.123"
        '
        'zvalue
        '
        Me.zvalue.Location = New System.Drawing.Point(64, 64)
        Me.zvalue.Name = "zvalue"
        Me.zvalue.Size = New System.Drawing.Size(112, 16)
        Me.zvalue.Text = "12345678.123"
        '
        'cfgstatus
        '
        Me.cfgstatus.Location = New System.Drawing.Point(8, 88)
        Me.cfgstatus.Name = "cfgstatus"
        Me.cfgstatus.Size = New System.Drawing.Size(224, 40)
        Me.cfgstatus.Text = "The cfgfile is"
        '
        'dbstatus
        '
        Me.dbstatus.Location = New System.Drawing.Point(8, 128)
        Me.dbstatus.Name = "dbstatus"
        Me.dbstatus.Size = New System.Drawing.Size(224, 40)
        Me.dbstatus.Text = "The db file is"
        '
        'recstatus
        '
        Me.recstatus.Location = New System.Drawing.Point(8, 168)
        Me.recstatus.Name = "recstatus"
        Me.recstatus.Size = New System.Drawing.Size(176, 16)
        Me.recstatus.Text = "Records in the database"
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.Add(Me.MenuItem1)
        '
        'MenuItem1
        '
        Me.MenuItem1.Text = "Continue"
        '
        'logstatus
        '
        Me.logstatus.Location = New System.Drawing.Point(8, 217)
        Me.logstatus.Name = "logstatus"
        Me.logstatus.Size = New System.Drawing.Size(224, 40)
        Me.logstatus.Text = "The log file is"
        '
        'frmBoot
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(240, 268)
        Me.ControlBox = False
        Me.Controls.Add(Me.logstatus)
        Me.Controls.Add(Me.recstatus)
        Me.Controls.Add(Me.dbstatus)
        Me.Controls.Add(Me.cfgstatus)
        Me.Controls.Add(Me.zvalue)
        Me.Controls.Add(Me.yvalue)
        Me.Controls.Add(Me.xvalue)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmBoot"
        Me.Text = "EDM-Mobile"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click

        Me.DialogResult = DialogResult.None

    End Sub
End Class
