Public Class frmStatus
    Inherits System.Windows.Forms.Form
    Friend WithEvents statustxt As System.Windows.Forms.TextBox

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
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.statustxt = New System.Windows.Forms.TextBox
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.SuspendLayout()
        '
        'statustxt
        '
        Me.statustxt.Location = New System.Drawing.Point(0, 0)
        Me.statustxt.Multiline = True
        Me.statustxt.Name = "statustxt"
        Me.statustxt.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.statustxt.Size = New System.Drawing.Size(240, 268)
        Me.statustxt.TabIndex = 0
        Me.statustxt.Text = "TextBox1"
        Me.statustxt.WordWrap = False
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.Add(Me.MenuItem3)
        '
        'MenuItem3
        '
        Me.MenuItem3.Text = "Close"
        '
        'frmStatus
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(240, 268)
        Me.ControlBox = False
        Me.Controls.Add(Me.statustxt)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmStatus"
        Me.Text = "Status"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmStatus_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub frmStatus_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated

        Dim status As Integer

        Call table_status(options(22), status)

        statustxt.Visible = False
        statustxt.Refresh()
        statustxt.Text = ""
        statustxt.Text = statustxt.Text + "Configuration file : " + cfgfile + vbCrLf
        statustxt.Text = statustxt.Text + "Total station type : " + options(4) + vbCrLf
        statustxt.Text = statustxt.Text + "Points table name  : " + options(22) + vbCrLf
        statustxt.Text = statustxt.Text + "Number of records  : " + Trim(CStr(status)) + vbCrLf
        statustxt.Text = statustxt.Text + "Database name      : " + options(1) + vbCrLf
        statustxt.Text = statustxt.Text + "COM Port settings  : " + options(2) + ":" + options(3) + vbCrLf
        statustxt.Text = statustxt.Text + "Current station name : " + currentstationname + vbCrLf
        statustxt.Text = statustxt.Text + "Current station X    : " + FormatNumber(currentstationx, 3) + vbCrLf
        statustxt.Text = statustxt.Text + "Current station Y    : " + FormatNumber(currentstationy, 3) + vbCrLf
        statustxt.Text = statustxt.Text + "Current station Z    : " + FormatNumber(currentstationz, 3) + vbCrLf
        If use_limitchecking Then
            statustxt.Text = statustxt.Text + "Limit checking       : Yes" + vbCrLf
        Else
            statustxt.Text = statustxt.Text + "Limit checking       : No" + vbCrLf
        End If
        statustxt.Text = statustxt.Text + "App.path : " + System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + vbCrLf
        statustxt.Text = statustxt.Text + "edmini : " + edmini + vbCrLf
        statustxt.Text = statustxt.Text + "fpath : " + fpath + vbCrLf
        If commcontrol.IsOpen Then
            statustxt.Text = statustxt.Text + "Comport : Open" + Chr(13)
        Else
            statustxt.Text = statustxt.Text + "Comport : Closed" + Chr(13)
        End If
        statustxt.Visible = True

    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click

        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub
End Class
