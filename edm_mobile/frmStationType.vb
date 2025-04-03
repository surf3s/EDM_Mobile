Public Class frmStationType
    Inherits System.Windows.Forms.Form
    Friend WithEvents topcon As System.Windows.Forms.RadioButton
    Friend WithEvents Wild As System.Windows.Forms.RadioButton
    Friend WithEvents Leica As System.Windows.Forms.RadioButton
    Friend WithEvents Sokkia As System.Windows.Forms.RadioButton
    Friend WithEvents Notcabled As System.Windows.Forms.RadioButton
    Friend WithEvents simulationmode As System.Windows.Forms.RadioButton
    Friend WithEvents Label1 As System.Windows.Forms.Label

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
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.topcon = New System.Windows.Forms.RadioButton
        Me.Wild = New System.Windows.Forms.RadioButton
        Me.Leica = New System.Windows.Forms.RadioButton
        Me.Sokkia = New System.Windows.Forms.RadioButton
        Me.Notcabled = New System.Windows.Forms.RadioButton
        Me.simulationmode = New System.Windows.Forms.RadioButton
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        '
        'topcon
        '
        Me.topcon.Location = New System.Drawing.Point(8, 32)
        Me.topcon.Size = New System.Drawing.Size(120, 16)
        Me.topcon.Text = "Topcon"
        '
        'Wild
        '
        Me.Wild.Location = New System.Drawing.Point(8, 56)
        Me.Wild.Size = New System.Drawing.Size(120, 16)
        Me.Wild.Text = "Wild"
        '
        'Leica
        '
        Me.Leica.Location = New System.Drawing.Point(8, 80)
        Me.Leica.Size = New System.Drawing.Size(120, 16)
        Me.Leica.Text = "Leica"
        '
        'Sokkia
        '
        Me.Sokkia.Location = New System.Drawing.Point(8, 104)
        Me.Sokkia.Size = New System.Drawing.Size(120, 16)
        Me.Sokkia.Text = "Sokkia"
        '
        'Notcabled
        '
        Me.Notcabled.Location = New System.Drawing.Point(8, 128)
        Me.Notcabled.Size = New System.Drawing.Size(120, 16)
        Me.Notcabled.Text = "Not cabled"
        '
        'simulationmode
        '
        Me.simulationmode.Location = New System.Drawing.Point(8, 152)
        Me.simulationmode.Size = New System.Drawing.Size(120, 16)
        Me.simulationmode.Text = "Simulation mode"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Size = New System.Drawing.Size(144, 16)
        Me.Label1.Text = "Select the total station  :"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular)
        Me.Label2.Location = New System.Drawing.Point(8, 176)
        Me.Label2.Size = New System.Drawing.Size(240, 56)
        Me.Label2.Text = "When the station is selected, this screen may pause for a long time while communi" & _
        "cation with the station is attempted.  If this fails, use Station Comport Settin" & _
        "g."
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(128, 40)
        Me.Label3.Size = New System.Drawing.Size(104, 128)
        Me.Label3.Text = "This program has been tested only with Leica and Topcon series stations. Previous" & _
        " versions were tested with Wild and Sokkia."
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.Add(Me.MenuItem3)
        '
        'MenuItem3
        '
        Me.MenuItem3.Text = "Close"
        '
        'frmStationType
        '
        Me.ControlBox = False
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.simulationmode)
        Me.Controls.Add(Me.Notcabled)
        Me.Controls.Add(Me.Sokkia)
        Me.Controls.Add(Me.Leica)
        Me.Controls.Add(Me.Wild)
        Me.Controls.Add(Me.topcon)
        Me.Menu = Me.MainMenu1
        Me.Text = "Total Station Type"

    End Sub

#End Region
    Dim inprocess As Boolean


    Private Sub topcon_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles topcon.CheckedChanged

        If inprocess Then Exit Sub
        Call putstationini("TOPCON")
        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub

    Private Sub Wild_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Wild.CheckedChanged

        If inprocess Then Exit Sub
        Call putstationini("WILD")
        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub

    Private Sub Sokkia_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Sokkia.CheckedChanged

        If inprocess Then Exit Sub
        Call putstationini("SOKKIA")
        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub

    Private Sub simulationmode_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles simulationmode.CheckedChanged

        If inprocess Then Exit Sub
        Call putstationini("SIMULATION")
        Call close_com()
        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub

    Private Sub Notcabled_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Notcabled.CheckedChanged

        If inprocess Then Exit Sub
        Call putstationini("NONE")
        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub

    Private Sub Leica_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Leica.CheckedChanged

        If inprocess Then Exit Sub
        Call putstationini("LEICA")
        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub

    Private Sub frmStationType_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated

        inprocess = True

        Select Case options(4)
            Case "TOPCON"
                topcon.Checked = True
            Case "WILD"
                Wild.Checked = True
            Case "LEICA"
                Leica.Checked = True
            Case "SOKKIA"
                Sokkia.Checked = True
            Case "SIMULATION"
                simulationmode.Checked = True
            Case "NONE"
                Notcabled.Checked = True
            Case Else
        End Select

        inprocess = False

    End Sub

    Private Sub putstationini(ByVal stationname As String)

        Dim status As Integer
        Dim iniclass As String
        Dim inidata(2, 2) As String

        Cursor.Current = Cursors.WaitCursor
        options(4) = stationname
        Call open_com()
        iniclass = "[EDM-MOBILE]"
        inidata(1, 1) = "Total Station"
        inidata(1, 2) = options(4)
        Call WriteIni(edmini, iniclass, inidata, status)
        Cursor.Current = Cursors.WaitCursor

    End Sub
    Private Sub frmStationType_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click

        Me.DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class
