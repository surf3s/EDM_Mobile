Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlServerCe
Imports System.io
Imports System
Public Class frmMain

    Inherits System.Windows.Forms.Form
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu

    Declare Function SetSystemPowerState Lib "Coredll" ( _
     ByVal psState As String, _
     ByVal StateFlags As Integer, _
     ByVal Options As Integer) As Integer

    Const POWER_STATE_ON As Integer = &H10000
    Const POWER_STATE_OFF As Integer = &H20000
    Const POWER_STATE_SUSPEND As Integer = &H200000
    Const POWER_FORCE As Integer = 4096
    Friend WithEvents menuspeedbuttons As System.Windows.Forms.MenuItem
    Friend WithEvents aboutus As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents barcode_scan As System.Windows.Forms.Panel
    Friend WithEvents unit_id As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Battery_test As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuViewPoints As System.Windows.Forms.MenuItem
    Const POWER_STATE_RESET As Integer = &H800000 'this wil make the device soft reset! :)

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
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem16 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem19 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem23 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem24 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem27 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem29 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem34 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem36 As System.Windows.Forms.MenuItem
    Friend WithEvents Menunewsite As System.Windows.Forms.MenuItem
    Friend WithEvents Menuopensite As System.Windows.Forms.MenuItem
    Friend WithEvents Menunewtable As System.Windows.Forms.MenuItem
    Friend WithEvents Menuopentable As System.Windows.Forms.MenuItem
    Friend WithEvents Menuexport As System.Windows.Forms.MenuItem
    Friend WithEvents Menuexit As System.Windows.Forms.MenuItem
    Friend WithEvents Menuprisms As System.Windows.Forms.MenuItem
    Friend WithEvents Menuunits As System.Windows.Forms.MenuItem
    Friend WithEvents Menudatums As System.Windows.Forms.MenuItem
    Friend WithEvents Menupoints As System.Windows.Forms.MenuItem
    Friend WithEvents Menuoptions As System.Windows.Forms.MenuItem
    Friend WithEvents Menupointsetup As System.Windows.Forms.MenuItem
    Friend WithEvents Menudeletelast As System.Windows.Forms.MenuItem
    Friend WithEvents Menueditlast As System.Windows.Forms.MenuItem
    Friend WithEvents Menuresetconnection As System.Windows.Forms.MenuItem
    Friend WithEvents Menucomport As System.Windows.Forms.MenuItem
    Friend WithEvents Menuadddatum As System.Windows.Forms.MenuItem
    Friend WithEvents Menuinitdirect As System.Windows.Forms.MenuItem
    Friend WithEvents Menuinitonepoint As System.Windows.Forms.MenuItem
    Friend WithEvents Menuinittwopoint As System.Windows.Forms.MenuItem
    Friend WithEvents Menuinitthreepoint As System.Windows.Forms.MenuItem
    Friend WithEvents Menuinitangleonly As System.Windows.Forms.MenuItem
    Friend WithEvents Menuverify As System.Windows.Forms.MenuItem
    Friend WithEvents Menucurrent As System.Windows.Forms.MenuItem
    Friend WithEvents Menuhelp As System.Windows.Forms.MenuItem
    Friend WithEvents Menustatus As System.Windows.Forms.MenuItem
    Friend WithEvents Menuabout As System.Windows.Forms.MenuItem
    Friend WithEvents Menureadme As System.Windows.Forms.MenuItem
    Friend WithEvents Menusleep As System.Windows.Forms.MenuItem
    Friend WithEvents ofd As System.Windows.Forms.OpenFileDialog
    Friend WithEvents sfd As System.Windows.Forms.SaveFileDialog
    Friend WithEvents Menudeleteseries As System.Windows.Forms.MenuItem
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents frame2 As System.Windows.Forms.Panel
    Friend WithEvents frame1 As System.Windows.Forms.Panel
    Friend WithEvents recordpoint As System.Windows.Forms.Button
    Friend WithEvents continuepoint As System.Windows.Forms.Button
    Friend WithEvents measurepoint As System.Windows.Forms.Button
    Friend WithEvents MenuViewEDMINI As System.Windows.Forms.MenuItem
    Friend WithEvents MenuViewCFG As System.Windows.Forms.MenuItem
    Private components As System.ComponentModel.IContainer
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents statusmessage As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.Menunewsite = New System.Windows.Forms.MenuItem
        Me.Menuopensite = New System.Windows.Forms.MenuItem
        Me.Menunewtable = New System.Windows.Forms.MenuItem
        Me.Menuopentable = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.Menuexport = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.Menuexit = New System.Windows.Forms.MenuItem
        Me.MenuItem10 = New System.Windows.Forms.MenuItem
        Me.Menuprisms = New System.Windows.Forms.MenuItem
        Me.Menuunits = New System.Windows.Forms.MenuItem
        Me.Menudatums = New System.Windows.Forms.MenuItem
        Me.Menupoints = New System.Windows.Forms.MenuItem
        Me.MenuItem16 = New System.Windows.Forms.MenuItem
        Me.Menuoptions = New System.Windows.Forms.MenuItem
        Me.menuspeedbuttons = New System.Windows.Forms.MenuItem
        Me.Menupointsetup = New System.Windows.Forms.MenuItem
        Me.MenuItem19 = New System.Windows.Forms.MenuItem
        Me.Menudeletelast = New System.Windows.Forms.MenuItem
        Me.Menudeleteseries = New System.Windows.Forms.MenuItem
        Me.Menueditlast = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.MenuViewPoints = New System.Windows.Forms.MenuItem
        Me.Menuresetconnection = New System.Windows.Forms.MenuItem
        Me.Battery_test = New System.Windows.Forms.MenuItem
        Me.MenuItem23 = New System.Windows.Forms.MenuItem
        Me.MenuItem24 = New System.Windows.Forms.MenuItem
        Me.Menucomport = New System.Windows.Forms.MenuItem
        Me.MenuItem27 = New System.Windows.Forms.MenuItem
        Me.Menuadddatum = New System.Windows.Forms.MenuItem
        Me.MenuItem29 = New System.Windows.Forms.MenuItem
        Me.Menuinitdirect = New System.Windows.Forms.MenuItem
        Me.Menuinitonepoint = New System.Windows.Forms.MenuItem
        Me.Menuinittwopoint = New System.Windows.Forms.MenuItem
        Me.Menuinitthreepoint = New System.Windows.Forms.MenuItem
        Me.MenuItem34 = New System.Windows.Forms.MenuItem
        Me.Menuinitangleonly = New System.Windows.Forms.MenuItem
        Me.MenuItem36 = New System.Windows.Forms.MenuItem
        Me.Menuverify = New System.Windows.Forms.MenuItem
        Me.Menucurrent = New System.Windows.Forms.MenuItem
        Me.Menuhelp = New System.Windows.Forms.MenuItem
        Me.Menuabout = New System.Windows.Forms.MenuItem
        Me.Menureadme = New System.Windows.Forms.MenuItem
        Me.MenuViewEDMINI = New System.Windows.Forms.MenuItem
        Me.MenuViewCFG = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.Menustatus = New System.Windows.Forms.MenuItem
        Me.Menusleep = New System.Windows.Forms.MenuItem
        Me.ofd = New System.Windows.Forms.OpenFileDialog
        Me.sfd = New System.Windows.Forms.SaveFileDialog
        Me.frame2 = New System.Windows.Forms.Panel
        Me.Button6 = New System.Windows.Forms.Button
        Me.Button5 = New System.Windows.Forms.Button
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.frame1 = New System.Windows.Forms.Panel
        Me.measurepoint = New System.Windows.Forms.Button
        Me.continuepoint = New System.Windows.Forms.Button
        Me.recordpoint = New System.Windows.Forms.Button
        Me.statusmessage = New System.Windows.Forms.Label
        Me.aboutus = New System.Windows.Forms.Panel
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.barcode_scan = New System.Windows.Forms.Panel
        Me.unit_id = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.frame2.SuspendLayout()
        Me.frame1.SuspendLayout()
        Me.aboutus.SuspendLayout()
        Me.barcode_scan.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.Add(Me.MenuItem1)
        Me.MainMenu1.MenuItems.Add(Me.MenuItem10)
        Me.MainMenu1.MenuItems.Add(Me.Menuresetconnection)
        Me.MainMenu1.MenuItems.Add(Me.Menuhelp)
        Me.MainMenu1.MenuItems.Add(Me.Menusleep)
        '
        'MenuItem1
        '
        Me.MenuItem1.MenuItems.Add(Me.Menunewsite)
        Me.MenuItem1.MenuItems.Add(Me.Menuopensite)
        Me.MenuItem1.MenuItems.Add(Me.Menunewtable)
        Me.MenuItem1.MenuItems.Add(Me.Menuopentable)
        Me.MenuItem1.MenuItems.Add(Me.MenuItem6)
        Me.MenuItem1.MenuItems.Add(Me.Menuexport)
        Me.MenuItem1.MenuItems.Add(Me.MenuItem8)
        Me.MenuItem1.MenuItems.Add(Me.MenuItem3)
        Me.MenuItem1.MenuItems.Add(Me.MenuItem4)
        Me.MenuItem1.MenuItems.Add(Me.Menuexit)
        Me.MenuItem1.Text = "File"
        '
        'Menunewsite
        '
        Me.Menunewsite.Text = "New Site File (CFG/MDB)"
        '
        'Menuopensite
        '
        Me.Menuopensite.Text = "Open Site File (CFG/MDB)"
        '
        'Menunewtable
        '
        Me.Menunewtable.Text = "New Points File (Table)"
        '
        'Menuopentable
        '
        Me.Menuopentable.Text = "Open Points File (Table)"
        '
        'MenuItem6
        '
        Me.MenuItem6.Text = "-"
        '
        'Menuexport
        '
        Me.Menuexport.Text = "Export ASCII"
        '
        'MenuItem8
        '
        Me.MenuItem8.Text = "-"
        '
        'MenuItem3
        '
        Me.MenuItem3.Text = "Clear Log File"
        '
        'MenuItem4
        '
        Me.MenuItem4.Text = "-"
        '
        'Menuexit
        '
        Me.Menuexit.Text = "Exit"
        '
        'MenuItem10
        '
        Me.MenuItem10.MenuItems.Add(Me.Menuprisms)
        Me.MenuItem10.MenuItems.Add(Me.Menuunits)
        Me.MenuItem10.MenuItems.Add(Me.Menudatums)
        Me.MenuItem10.MenuItems.Add(Me.Menupoints)
        Me.MenuItem10.MenuItems.Add(Me.MenuItem16)
        Me.MenuItem10.MenuItems.Add(Me.Menuoptions)
        Me.MenuItem10.MenuItems.Add(Me.menuspeedbuttons)
        Me.MenuItem10.MenuItems.Add(Me.Menupointsetup)
        Me.MenuItem10.MenuItems.Add(Me.MenuItem19)
        Me.MenuItem10.MenuItems.Add(Me.Menudeletelast)
        Me.MenuItem10.MenuItems.Add(Me.Menudeleteseries)
        Me.MenuItem10.MenuItems.Add(Me.Menueditlast)
        Me.MenuItem10.MenuItems.Add(Me.MenuItem7)
        Me.MenuItem10.MenuItems.Add(Me.MenuViewPoints)
        Me.MenuItem10.Text = "Edit"
        '
        'Menuprisms
        '
        Me.Menuprisms.Text = "Prisms"
        '
        'Menuunits
        '
        Me.Menuunits.Text = "Excavation Units"
        '
        'Menudatums
        '
        Me.Menudatums.Text = "Datums"
        '
        'Menupoints
        '
        Me.Menupoints.Text = "Points"
        '
        'MenuItem16
        '
        Me.MenuItem16.Text = "-"
        '
        'Menuoptions
        '
        Me.Menuoptions.Text = "EDM-Mobile Options"
        '
        'menuspeedbuttons
        '
        Me.menuspeedbuttons.Text = "Speed Buttons"
        '
        'Menupointsetup
        '
        Me.Menupointsetup.Text = "Point Fields Setup"
        '
        'MenuItem19
        '
        Me.MenuItem19.Text = "-"
        '
        'Menudeletelast
        '
        Me.Menudeletelast.Text = "Delete Last Point"
        '
        'Menudeleteseries
        '
        Me.Menudeleteseries.Text = "Delete Last Series"
        '
        'Menueditlast
        '
        Me.Menueditlast.Text = "Edit Last Point"
        '
        'MenuItem7
        '
        Me.MenuItem7.Text = "-"
        '
        'MenuViewPoints
        '
        Me.MenuViewPoints.Text = "View Points"
        '
        'Menuresetconnection
        '
        Me.Menuresetconnection.MenuItems.Add(Me.Battery_test)
        Me.Menuresetconnection.MenuItems.Add(Me.MenuItem23)
        Me.Menuresetconnection.MenuItems.Add(Me.MenuItem24)
        Me.Menuresetconnection.MenuItems.Add(Me.Menucomport)
        Me.Menuresetconnection.MenuItems.Add(Me.MenuItem27)
        Me.Menuresetconnection.MenuItems.Add(Me.Menuadddatum)
        Me.Menuresetconnection.MenuItems.Add(Me.MenuItem29)
        Me.Menuresetconnection.MenuItems.Add(Me.Menuinitdirect)
        Me.Menuresetconnection.MenuItems.Add(Me.Menuinitonepoint)
        Me.Menuresetconnection.MenuItems.Add(Me.Menuinittwopoint)
        Me.Menuresetconnection.MenuItems.Add(Me.Menuinitthreepoint)
        Me.Menuresetconnection.MenuItems.Add(Me.MenuItem34)
        Me.Menuresetconnection.MenuItems.Add(Me.Menuinitangleonly)
        Me.Menuresetconnection.MenuItems.Add(Me.MenuItem36)
        Me.Menuresetconnection.MenuItems.Add(Me.Menuverify)
        Me.Menuresetconnection.MenuItems.Add(Me.Menucurrent)
        Me.Menuresetconnection.Text = "Station"
        '
        'Battery_test
        '
        Me.Battery_test.Text = "Battery Test"
        '
        'MenuItem23
        '
        Me.MenuItem23.Text = "Reset Connection"
        '
        'MenuItem24
        '
        Me.MenuItem24.Text = "-"
        '
        'Menucomport
        '
        Me.Menucomport.Text = "Settings"
        '
        'MenuItem27
        '
        Me.MenuItem27.Text = "-"
        '
        'Menuadddatum
        '
        Me.Menuadddatum.Text = "Add a Datum"
        '
        'MenuItem29
        '
        Me.MenuItem29.Text = "-"
        '
        'Menuinitdirect
        '
        Me.Menuinitdirect.Text = "Initialize Direct"
        '
        'Menuinitonepoint
        '
        Me.Menuinitonepoint.Text = "Initialize One Point"
        '
        'Menuinittwopoint
        '
        Me.Menuinittwopoint.Text = "Initialize Two Points"
        '
        'Menuinitthreepoint
        '
        Me.Menuinitthreepoint.Text = "Initialize Three Points"
        '
        'MenuItem34
        '
        Me.MenuItem34.Text = "-"
        '
        'Menuinitangleonly
        '
        Me.Menuinitangleonly.Text = "Initialize Angle Only"
        '
        'MenuItem36
        '
        Me.MenuItem36.Text = "-"
        '
        'Menuverify
        '
        Me.Menuverify.Text = "Verify"
        '
        'Menucurrent
        '
        Me.Menucurrent.Text = "Current Coordinates"
        '
        'Menuhelp
        '
        Me.Menuhelp.MenuItems.Add(Me.Menuabout)
        Me.Menuhelp.MenuItems.Add(Me.Menureadme)
        Me.Menuhelp.MenuItems.Add(Me.MenuViewEDMINI)
        Me.Menuhelp.MenuItems.Add(Me.MenuViewCFG)
        Me.Menuhelp.MenuItems.Add(Me.MenuItem5)
        Me.Menuhelp.MenuItems.Add(Me.MenuItem2)
        Me.Menuhelp.MenuItems.Add(Me.Menustatus)
        Me.Menuhelp.Text = "Help"
        '
        'Menuabout
        '
        Me.Menuabout.Text = "About"
        '
        'Menureadme
        '
        Me.Menureadme.Text = "Read Me"
        '
        'MenuViewEDMINI
        '
        Me.MenuViewEDMINI.Text = "View EDM.INI"
        '
        'MenuViewCFG
        '
        Me.MenuViewCFG.Text = "View CFG File"
        '
        'MenuItem5
        '
        Me.MenuItem5.Text = "View Log File"
        '
        'MenuItem2
        '
        Me.MenuItem2.Text = "Database Tables"
        '
        'Menustatus
        '
        Me.Menustatus.Text = "Status"
        '
        'Menusleep
        '
        Me.Menusleep.Text = "Sleep"
        '
        'frame2
        '
        Me.frame2.Controls.Add(Me.Button6)
        Me.frame2.Controls.Add(Me.Button5)
        Me.frame2.Controls.Add(Me.Button4)
        Me.frame2.Controls.Add(Me.Button3)
        Me.frame2.Controls.Add(Me.Button2)
        Me.frame2.Controls.Add(Me.Button1)
        Me.frame2.Location = New System.Drawing.Point(8, 24)
        Me.frame2.Name = "frame2"
        Me.frame2.Size = New System.Drawing.Size(224, 160)
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(160, 88)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(64, 64)
        Me.Button6.TabIndex = 8
        Me.Button6.Text = "Button6"
        Me.Button6.Visible = False
        '
        'Button5
        '
        Me.Button5.Location = New System.Drawing.Point(80, 88)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(64, 64)
        Me.Button5.TabIndex = 7
        Me.Button5.Text = "Button5"
        Me.Button5.Visible = False
        '
        'Button4
        '
        Me.Button4.Location = New System.Drawing.Point(0, 88)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(64, 64)
        Me.Button4.TabIndex = 6
        Me.Button4.Text = "Button4"
        Me.Button4.Visible = False
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(160, 8)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(64, 64)
        Me.Button3.TabIndex = 5
        Me.Button3.Text = "Button3"
        Me.Button3.Visible = False
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(80, 8)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(64, 64)
        Me.Button2.TabIndex = 4
        Me.Button2.Text = "Button2"
        Me.Button2.Visible = False
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(0, 8)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(64, 64)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Button1"
        Me.Button1.Visible = False
        '
        'frame1
        '
        Me.frame1.Controls.Add(Me.measurepoint)
        Me.frame1.Controls.Add(Me.continuepoint)
        Me.frame1.Controls.Add(Me.recordpoint)
        Me.frame1.Location = New System.Drawing.Point(8, 192)
        Me.frame1.Name = "frame1"
        Me.frame1.Size = New System.Drawing.Size(224, 72)
        '
        'measurepoint
        '
        Me.measurepoint.Enabled = False
        Me.measurepoint.Location = New System.Drawing.Point(160, 8)
        Me.measurepoint.Name = "measurepoint"
        Me.measurepoint.Size = New System.Drawing.Size(64, 48)
        Me.measurepoint.TabIndex = 2
        Me.measurepoint.Text = "Measure"
        '
        'continuepoint
        '
        Me.continuepoint.Enabled = False
        Me.continuepoint.Location = New System.Drawing.Point(80, 8)
        Me.continuepoint.Name = "continuepoint"
        Me.continuepoint.Size = New System.Drawing.Size(64, 48)
        Me.continuepoint.TabIndex = 1
        Me.continuepoint.Text = "Continue"
        '
        'recordpoint
        '
        Me.recordpoint.Enabled = False
        Me.recordpoint.Location = New System.Drawing.Point(0, 8)
        Me.recordpoint.Name = "recordpoint"
        Me.recordpoint.Size = New System.Drawing.Size(64, 48)
        Me.recordpoint.TabIndex = 0
        Me.recordpoint.Text = "Record"
        '
        'statusmessage
        '
        Me.statusmessage.Location = New System.Drawing.Point(8, 8)
        Me.statusmessage.Name = "statusmessage"
        Me.statusmessage.Size = New System.Drawing.Size(224, 16)
        Me.statusmessage.Text = "Last Point Recorded"
        Me.statusmessage.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.statusmessage.Visible = False
        '
        'aboutus
        '
        Me.aboutus.Controls.Add(Me.Label5)
        Me.aboutus.Controls.Add(Me.Label4)
        Me.aboutus.Controls.Add(Me.Label3)
        Me.aboutus.Controls.Add(Me.Label1)
        Me.aboutus.Location = New System.Drawing.Point(11, 8)
        Me.aboutus.Name = "aboutus"
        Me.aboutus.Size = New System.Drawing.Size(218, 186)
        Me.aboutus.Visible = False
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(3, 142)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(208, 46)
        Me.Label5.Text = "For program updates and info - see www.oldstoneage.com/software"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(5, 74)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(208, 16)
        Me.Label4.Text = "An OldStoneAge.Com Production"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(34, 34)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(151, 30)
        Me.Label3.Text = "By Shannon P. McPherron and Harold L. Dibble"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Bold)
        Me.Label1.Location = New System.Drawing.Point(35, 5)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(151, 19)
        Me.Label1.Text = "EDM-Mobile 1.10"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'barcode_scan
        '
        Me.barcode_scan.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.barcode_scan.Controls.Add(Me.unit_id)
        Me.barcode_scan.Controls.Add(Me.Label2)
        Me.barcode_scan.Location = New System.Drawing.Point(11, 35)
        Me.barcode_scan.Name = "barcode_scan"
        Me.barcode_scan.Size = New System.Drawing.Size(218, 141)
        Me.barcode_scan.Visible = False
        '
        'unit_id
        '
        Me.unit_id.Font = New System.Drawing.Font("Tahoma", 21.0!, System.Drawing.FontStyle.Regular)
        Me.unit_id.Location = New System.Drawing.Point(5, 40)
        Me.unit_id.Name = "unit_id"
        Me.unit_id.Size = New System.Drawing.Size(206, 40)
        Me.unit_id.TabIndex = 1
        Me.unit_id.Text = "IRHOUD-88888"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(5, 20)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(142, 15)
        Me.Label2.Text = "Next Unit-ID"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(240, 268)
        Me.ControlBox = False
        Me.Controls.Add(Me.barcode_scan)
        Me.Controls.Add(Me.aboutus)
        Me.Controls.Add(Me.frame2)
        Me.Controls.Add(Me.statusmessage)
        Me.Controls.Add(Me.frame1)
        Me.MaximizeBox = False
        Me.Menu = Me.MainMenu1
        Me.Name = "frmMain"
        Me.Text = "EDM-Mobile"
        Me.frame2.ResumeLayout(False)
        Me.frame1.ResumeLayout(False)
        Me.aboutus.ResumeLayout(False)
        Me.barcode_scan.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Shared Sub Main()

        Application.Run(New frmMain)

    End Sub

    Private Sub frmMain_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated

        If returnedstatus <> "" Then
            statusmessage.Text = returnedstatus
            statusmessage.Visible = True
        Else
            statusmessage.Visible = False
        End If

        If bar_code_pre_scan Then
            barcode_scan.Visible = True
            frame2.Visible = False
            unit_id.Focus()
            unit_id.SelectAll()
        Else
            barcode_scan.Visible = False
            frame2.Visible = True
        End If

    End Sub

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim answer As Integer
        Dim status As Integer
        Dim needtoconvert As Integer = 0
        Dim backup As String

        aboutus.Visible = True
        Me.Show()
        Cursor.Current = Cursors.WaitCursor
        MenuItem1.Enabled = False
        MenuItem10.Enabled = False
        Menuresetconnection.Enabled = False
        Menuhelp.Enabled = False
        Menusleep.Enabled = False
        Me.Refresh()
        Randomize()

        'do this only when debugging on PC
        in_emulation = True

        'Begin with an empty edm station and setup
        currentstationname = ""
        currentstationx = 0
        currentstationy = 0
        currentstationz = 0

        'Setup the default path
        If fpath = "" Then
            If Not Directory.Exists("\my documents") Then Directory.CreateDirectory("\My Documents")
            fpath = "\My Documents\"
        End If

        If edmini = "" Then edmini = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) + "\edm.ini"

        tablesopen = False

        If Not File.Exists(edmini) Then
            'Offer to use a default configuation
            answer = MsgBox("Do you want to begin the program with the default configuration that will allow you to begin recording points?", MsgBoxStyle.YesNoCancel Or MsgBoxStyle.Question, "EDM-Mobile")
            If answer = 6 Then
                options(1) = "edm"
                options(22) = "edm"
                options(2) = "COM1"
                options(3) = "1200,E,7,1"
                options(4) = "SIMULATION"
                options(11) = "CE"
                options(12) = "no"
                options(13) = "Stadia"
                options(14) = "2"
                options(15) = "0"
                options(17) = "Sitename"
                options(19) = "no"
                Call use_defaults()
                Call write_edmce_ini(edmini, status)
                MsgBox("EDM-Mobile is currently in a simulation mode. If you are connected to a total station, you need to select the total station type and communication settings (under Station menu).", MsgBoxStyle.Information, "EDM-Mobile")

            ElseIf answer = 7 Then
                MsgBox("Before you can use this program, you will need to create or open a configuration file.", MsgBoxStyle.Information, "EDM-Mobile")

            Else
                Application.Exit()

            End If

        Else
            'read info from last time program was run
            Dim inidata(10, 2) As String
            Dim iniclass As String

            iniclass = "[EDM-MOBILE]"
            inidata(1, 1) = "CFG_File"
            inidata(2, 1) = "COM"
            inidata(3, 1) = "Settings"
            inidata(4, 1) = "Total Station"
            Call ReadIni(edmini, iniclass, inidata, status)
            cfgfile = inidata(1, 2)
            If inidata(2, 2) <> "" Then options(2) = inidata(2, 2)
            If inidata(3, 2) <> "" Then options(3) = inidata(3, 2)
            If inidata(4, 2) <> "" Then options(4) = inidata(4, 2)

            If cfgfile <> "" Then

                If Not File.Exists(cfgfile) Then
                    cfgfile = ""
                    MsgBox("The file " + cfgfile + " could not be found.  Before you can use the program, you will need to open a configuration file and create a default configuration.", MsgBoxStyle.Information, "EDM-Mobile")

                Else
                    Call parsecfg(needtoconvert)
                    Call fix_cfg_info()
                    Call setup_dbs()
                    Call setup_tbs()
                    Call check_field_sizes()

                    If needtoconvert = 1 Then
                        'make a backup of the existing one
                        backup = leftstr(cfgfile, InStr(cfgfile, ".") - 1) + "-OLD.cfg"
                        If File.Exists(backup) Then File.Delete(backup)
                        File.Copy(cfgfile, backup)
                        Call write_datafile_ini(cfgfile, status)
                    End If

                    Call setup_buttons()

                    inidata(2, 1) = "Total Station"
                    inidata(3, 1) = "COM"
                    inidata(4, 1) = "Settings"
                    inidata(5, 1) = "CurrentStation"
                    inidata(6, 1) = "CurrentStationX"
                    inidata(7, 1) = "CurrentStationY"
                    inidata(8, 1) = "CurrentStationZ"
                    inidata(9, 1) = "ReferenceDatum"
                    inidata(10, 1) = "ReferenceDatum2"
                    iniclass = "[EDM]"
                    Call ReadIni(cfgfile, iniclass, inidata, status)
                    If inidata(9, 2) <> "" Then referencedatum = inidata(9, 2)
                    If inidata(10, 2) <> "" Then referencedatum2 = inidata(10, 2)
                    If inidata(5, 2) <> "" Then currentstationname = inidata(5, 2)
                    If inidata(6, 2) <> "" Then currentstationx = CDbl(inidata(6, 2))
                    If inidata(7, 2) <> "" Then currentstationy = CDbl(inidata(7, 2))
                    If inidata(8, 2) <> "" Then currentstationz = CDbl(inidata(8, 2))
                    If inidata(3, 2) <> "" And options(2) = "" Then options(2) = inidata(3, 2)
                    If inidata(4, 2) <> "" And options(3) = "" Then options(3) = inidata(4, 2)
                    If inidata(2, 2) <> "" And options(4) = "" Then options(4) = inidata(2, 2)
                    Select Case options(4)
                        Case "TOPCON", "WILD", "SOKKIA", "LEICA", "LEICA-GEO"
                            Call open_com()
                        Case Else
                    End Select

                End If
                If cfgfile <> "" Then
                    LogFile = Path.GetDirectoryName(cfgfile) + "\" + Path.GetFileNameWithoutExtension(cfgfile) + "_log.txt"
                End If
            End If

        End If

        'Now put some status information on the screen
        Dim frmboot As New frmBoot
        frmboot.xvalue.Text = FormatNumber(currentstationx, 3)
        frmboot.yvalue.Text = FormatNumber(currentstationy, 3)
        frmboot.zvalue.Text = FormatNumber(currentstationz, 3)
        Call table_status(options(22), status)
        frmboot.recstatus.Text = LTrim(CStr(status)) & " records in the data table " & options(22)
        frmboot.cfgstatus.Text = "CFG file : " & cfgfile
        frmboot.dbstatus.Text = "Database file : " & options(1)
        If WriteLogFile Then
            frmboot.logstatus.Text = "Log file : " & LogFile
        Else
            frmboot.logstatus.Text = "Log file not in use."
        End If
        aboutus.Visible = False
        Cursor.Current = Cursors.Default
        frmboot.ShowDialog()

        frame1.Left = CInt(Me.Width / 2 - frame1.Width / 2)
        frame2.Left = CInt(Me.Width / 2 - frame2.Width / 2)

        back_to_edit = False
        back_to_main = False
        refresh_point_grid = True

        recordpoint.Enabled = True
        continuepoint.Enabled = True
        measurepoint.Enabled = True
        MenuItem1.Enabled = True
        MenuItem10.Enabled = True
        Menuresetconnection.Enabled = True
        Menuhelp.Enabled = True
        Menusleep.Enabled = True

    End Sub

    Private Sub Menuexit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Menuexit.Click

        Application.Exit()

    End Sub

    Private Sub Menuopensite_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Menuopensite.Click

        ofd.Filter = "*.cfg|*.cfg"
        ofd.ShowDialog()
        If ofd.FileName <> "" Then
            Cursor.Current = Cursors.WaitCursor

            Dim a As Integer
            Dim needtoconvert As Integer
            Dim backup As String
            Dim inidata(100, 2) As String
            Dim iniclass As String
            Dim wstatus As Integer
            Dim status As Integer
            Dim dbname As String

            cfgfile = ofd.FileName
            fpath = Path.GetFullPath(cfgfile)

            For a = 1 To vars
                vtype(a) = 0
                vardefault(a) = ""
                varprompt(a) = ""
                varlist(a) = ""
                vlen(a) = 0
                vmenu(a) = ""
                vunique(a) = False
                vcarry(a) = False
                vincr(a) = False
                varprint(a) = False
            Next a

            Call parsecfg(needtoconvert)

            Call fix_cfg_info()

            dbname = options(1)
            If dbname = "" Then
                dbname = Path.GetDirectoryName(cfgfile) + "\" + Path.GetFileNameWithoutExtension(cfgfile) + ".sdf"
            ElseIf Not File.Exists(dbname) Then
                dbname = Path.GetDirectoryName(cfgfile) + "\" + Path.GetFileNameWithoutExtension(cfgfile) + ".sdf"
            End If
            options(1) = dbname

            Call setup_dbs()

            Call setup_tbs()

            Call setup_buttons()

            If needtoconvert = 1 Then
                'make a backup of the existing one
                backup = leftstr(cfgfile, InStr(cfgfile, ".") - 1) + "-OLD.cfg"
                If File.Exists(backup) Then File.Delete(backup)
                File.Copy(cfgfile, backup)
                Call write_datafile_ini(cfgfile, status)
            End If

            Call write_edmce_ini(edmini, status)

            iniclass = "[EDM]"
            inidata(2, 1) = "Database"
            inidata(2, 2) = options(1)
            inidata(3, 1) = "PointTable"
            inidata(3, 2) = options(22)
            Call WriteIni(cfgfile, iniclass, inidata, wstatus)

            refresh_point_grid = True

            If cfgfile <> "" Then
                LogFile = Path.GetDirectoryName(cfgfile) + "\" + Path.GetFileNameWithoutExtension(cfgfile) + "_log.txt"
            End If

            Cursor.Current = Cursors.Default
        End If

    End Sub

    Private Sub Menunewsite_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Menunewsite.Click

        Dim answer As Integer
        Dim makedefault As Boolean
        Dim old_dbname As String = options(1)

        makedefault = False
        If cfgfile <> "" Then
            answer = MsgBox("Create new site configuration using current configuration?", MsgBoxStyle.Question Or MsgBoxStyle.YesNoCancel)
            If answer = MsgBoxResult.Cancel Then
                Exit Sub
            ElseIf answer = MsgBoxResult.No Then
                MsgBox("A default configuration file will be created.", MsgBoxStyle.Information Or MsgBoxStyle.OkOnly)
                makedefault = True
            End If
        Else
            makedefault = True
        End If

        sfd.ShowDialog()

        If sfd.FileName <> "" Then

            Cursor.Current = Cursors.WaitCursor

            Dim n As String
            Dim a As Integer
            Dim dbname As String
            Dim status As Integer
            Dim flag As Integer
            Dim iniclass As String
            Dim inidata(2, 2) As String
            Dim prismdata(40, 3) As String

            n = sfd.FileName
            If InStr(n, ".") <> 0 Then n = leftstr(n, InStr(n, ".") - 1)
            flag = 1
            If File.Exists(n + ".sdf") Then
                Cursor.Current = Cursors.Default
                a = MsgBox("The file " + n + ".sdf" + " already exists.  Overwrite?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "EDM-Mobile")
                If a <> 6 Then Exit Sub
                Cursor.Current = Cursors.WaitCursor
                If a <> 6 Then flag = 0
            End If
            fpath = Path.GetFullPath(n)
            dbname = n + ".sdf"

            If makedefault And flag = 1 Then

                cfgfile = n & ".cfg"
                options(1) = dbname
                options(22) = Path.GetFileNameWithoutExtension(dbname)
                Call use_defaults()
                iniclass = "[EDM]"
                inidata(1, 1) = "DATABASE"
                inidata(1, 2) = options(1)
                inidata(2, 1) = "POINTTABLE"
                inidata(2, 2) = options(22)
                Call WriteIni(cfgfile, iniclass, inidata, status)
                iniclass = "[EDM-MOBILE]"
                inidata(1, 1) = "CFG_FILE"
                inidata(1, 2) = cfgfile
                Call WriteIni(edmini, iniclass, inidata, status)

            Else

                If InStr(dbname, "\") = 0 Then
                    dbname = fpath + dbname
                End If

                If InStr(dbname, ".") = 0 Then
                    dbname = dbname + ".sdf"
                End If

                cfgfile = n & ".cfg"
                Call write_ini_header(cfgfile)
                options(1) = dbname
                options(22) = Path.GetFileNameWithoutExtension(dbname)
                iniclass = "[EDM]"
                inidata(1, 1) = "DATABASE"
                inidata(1, 2) = options(1)
                inidata(2, 1) = "POINTTABLE"
                inidata(2, 2) = options(22)
                Call WriteIni(cfgfile, iniclass, inidata, status)
                iniclass = "[EDM-MOBILE]"
                inidata(1, 1) = "CFG_FILE"
                inidata(1, 2) = cfgfile
                Call WriteIni(edmini, iniclass, inidata, status)

                Call setup_dbs()

                If Me.Button1.Visible = True Then Call save_button(1, Me.Button1.Text, cfgfile)
                If Me.Button2.Visible = True Then Call save_button(2, Me.Button2.Text, cfgfile)
                If Me.Button3.Visible = True Then Call save_button(3, Me.Button3.Text, cfgfile)
                If Me.Button4.Visible = True Then Call save_button(4, Me.Button4.Text, cfgfile)
                If Me.Button5.Visible = True Then Call save_button(5, Me.Button5.Text, cfgfile)
                If Me.Button6.Visible = True Then Call save_button(6, Me.Button6.Text, cfgfile)

                'Now write section for each field
                For a = 1 To vars
                    Call savevar(a, cfgfile)
                Next a

            End If

            ' Move over prisms, datums, and units
            If old_dbname <> "" Then
                Dim myOldConnection As New SqlCeConnection("Datasource='" + old_dbname + "'")
                Dim cmdOld As SqlCeCommand = myOldConnection.CreateCommand
                myOldConnection.Open()

                Dim myNewConnection As New SqlCeConnection("Datasource='" + options(1) + "'")
                Dim cmdNew As SqlCeCommand = myNewConnection.CreateCommand
                myNewConnection.Open()

                cmdOld.CommandText = "SELECT * FROM edm_poles;"
                Dim rdr As SqlCeDataReader = cmdOld.ExecuteReader

                Dim f As Integer = 0
                Dim t As String = ""
                Dim field_names As String = ""
                For f = 0 To rdr.FieldCount - 1
                    If field_names = "" Then
                        field_names = rdr.GetName(f)
                    Else
                        field_names += "," + rdr.GetName(f)
                    End If
                Next

                Do While rdr.Read
                    t = ""
                    For f = 0 To rdr.FieldCount - 1
                        Select Case rdr.GetFieldType(f).ToString
                            Case "System.String"
                                If t = "" Then
                                    t = "'" + rdr.GetValue(f).ToString + "'"
                                Else
                                    t += ",'" + rdr.GetValue(f).ToString + "'"
                                End If
                            Case Else
                                If rdr.GetValue(f).ToString = "" Then
                                    If t = "" Then
                                        t = "0"
                                    Else
                                        t += ",0"
                                    End If
                                Else
                                    If t = "" Then
                                        t = rdr.GetValue(f).ToString
                                    Else
                                        t += "," + rdr.GetValue(f).ToString
                                    End If
                                End If
                        End Select
                    Next
                    cmdNew.CommandText = "INSERT INTO edm_poles (" + field_names + ") VALUES (" + t + ")"
                    cmdNew.ExecuteNonQuery()
                Loop
                rdr.Close()

                ' Now do the same for edm_units
                cmdOld.CommandText = "SELECT * FROM edm_units;"
                rdr = cmdOld.ExecuteReader

                field_names = ""
                For f = 0 To rdr.FieldCount - 1
                    If field_names = "" Then
                        field_names = rdr.GetName(f)
                    Else
                        field_names += "," + rdr.GetName(f)
                    End If
                Next

                Do While rdr.Read
                    t = ""
                    For f = 0 To rdr.FieldCount - 1
                        Select Case rdr.GetFieldType(f).ToString
                            Case "System.String"
                                If t = "" Then
                                    t = "'" + rdr.GetValue(f).ToString + "'"
                                Else
                                    t += ",'" + rdr.GetValue(f).ToString + "'"
                                End If
                            Case Else
                                If rdr.GetValue(f).ToString = "" Then
                                    If t = "" Then
                                        t = "0"
                                    Else
                                        t += ",0"
                                    End If
                                Else
                                    If t = "" Then
                                        t = rdr.GetValue(f).ToString
                                    Else
                                        t += "," + rdr.GetValue(f).ToString
                                    End If
                                End If
                        End Select
                    Next
                    cmdNew.CommandText = "INSERT INTO edm_units (" + field_names + ") VALUES (" + t + ")"
                    cmdNew.ExecuteNonQuery()
                Loop
                rdr.Close()

                ' Now do the same for edm_datums
                cmdOld.CommandText = "SELECT * FROM edm_datums;"
                rdr = cmdOld.ExecuteReader

                field_names = ""
                For f = 0 To rdr.FieldCount - 1
                    If field_names = "" Then
                        field_names = rdr.GetName(f)
                    Else
                        field_names += "," + rdr.GetName(f)
                    End If
                Next

                Do While rdr.Read
                    t = ""
                    For f = 0 To rdr.FieldCount - 1
                        Select Case rdr.GetFieldType(f).ToString
                            Case "System.String"
                                If t = "" Then
                                    t = "'" + rdr.GetValue(f).ToString + "'"
                                Else
                                    t += ",'" + rdr.GetValue(f).ToString + "'"
                                End If
                            Case "System.Date", "System.DateTime"
                                If t = "" Then
                                    t = "'" + rdr.GetValue(f).ToString + "'"
                                Else
                                    t += ",'" + rdr.GetValue(f).ToString + "'"
                                End If
                            Case Else
                                If rdr.GetValue(f).ToString = "" Then
                                    If t = "" Then
                                        t = "0"
                                    Else
                                        t += ",0"
                                    End If
                                Else
                                    If t = "" Then
                                        t = rdr.GetValue(f).ToString
                                    Else
                                        t += "," + rdr.GetValue(f).ToString
                                    End If
                                End If
                        End Select
                    Next
                    cmdNew.CommandText = "INSERT INTO edm_datums (" + field_names + ") VALUES (" + t + ")"
                    cmdNew.ExecuteNonQuery()
                Loop
                rdr.Close()

                myNewConnection.Close()
                myOldConnection.Close()

            End If

            If cfgfile <> "" Then
                LogFile = Path.GetDirectoryName(cfgfile) + "\" + Path.GetFileNameWithoutExtension(cfgfile) + "_log.txt"
            End If

            Cursor.Current = Cursors.Default

        End If

    End Sub

    Private Sub Menunewtable_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Menunewtable.Click
        Cursor.Current = Cursors.WaitCursor
        Dim form1 As New frmNewTable
        form1.ShowDialog()
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub Menuopentable_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Menuopentable.Click
        Cursor.Current = Cursors.WaitCursor
        Dim form1 As New frmOpenTable
        form1.ShowDialog()
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub Menuprisms_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Menuprisms.Click
        Cursor.Current = Cursors.WaitCursor
        Dim form1 As New frmPrisms
        form1.ShowDialog()
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub Menuunits_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Menuunits.Click
        Cursor.Current = Cursors.WaitCursor
        Dim form1 As New frmUnits
        form1.ShowDialog()
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub Menupoints_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Menupoints.Click
        Cursor.Current = Cursors.WaitCursor
        Dim form1 As New frmPoints
        form1.ShowDialog()
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub Menuoptions_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Menuoptions.Click
        Cursor.Current = Cursors.WaitCursor
        Dim form1 As New frmOptions
        form1.ShowDialog()
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub Menupointsetup_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Menupointsetup.Click
        Cursor.Current = Cursors.WaitCursor
        Dim form1 As New frmFieldSetup
        form1.ShowDialog()
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub Menudeletelast_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Menudeletelast.Click

        Dim points As New SqlCeConnection("datasource='" + options(1) + "'")
        Dim sql As SqlCeCommand = points.CreateCommand
        Cursor.Current = Cursors.WaitCursor
        sql.CommandText = "SELECT * FROM [" & options(22) & "] ORDER BY Recno Desc"
        points.Open()
        Dim rdr As SqlCeDataReader = sql.ExecuteReader
        If rdr.Read Then
            Dim unit As String = rdr.Item("unit").ToString
            Dim id As String = rdr.Item("ID").ToString
            Dim suffix As String = rdr.Item("SUFFIX").ToString
            Dim sqidsuffix As String = unit + "-" + id + "(" + suffix + ")"
            If MsgBox("Delete record with ID of " + sqidsuffix + " ?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                Dim sqldel As SqlCeCommand = points.CreateCommand
                sqldel.CommandText = "DELETE FROM [" & options(22) & "] WHERE recno=" + rdr.Item("recno").ToString
                sqldel.ExecuteNonQuery()

                'See if ID deleted is current ID in units file and if so back off one
                If IsNumeric(id) And CInt(suffix) = 0 Then
                    Dim r As Object
                    sql.CommandText = "SELECT ID FROM edm_units WHERE unit='" + unit + "'"
                    r = sql.ExecuteScalar
                    If IsNumeric(r) Then
                        If CInt(id) = CInt(r) Then
                            id = CStr(CInt(id) - 1)
                            sql.CommandText = "UPDATE edm_units SET id='" + id + "' WHERE unit='" + unit + "'"
                            sql.ExecuteNonQuery()
                        End If
                    End If
                End If
            End If
        Else
            MsgBox("The points table is empty.", MsgBoxStyle.Information)
        End If
        points.Close()
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub Menudeleteseries_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Menudeleteseries.Click

        Dim points As New SqlCeConnection("datasource='" + options(1) + "'")
        Dim sql As SqlCeCommand = points.CreateCommand
        Cursor.Current = Cursors.WaitCursor
        sql.CommandText = "SELECT * FROM [" & options(22) & "] ORDER BY Recno Desc"
        points.Open()
        Dim rdr As SqlCeDataReader = sql.ExecuteReader
        If rdr.Read Then
            Dim unit As String = rdr.Item("unit").ToString
            Dim id As String = rdr.Item("id").ToString
            Dim sqid As String = unit + "-" + id
            If MsgBox("Delete all records associated with an ID of " + sqid + " ?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                Dim sqldel As SqlCeCommand = points.CreateCommand
                sqldel.CommandText = "DELETE FROM [" & options(22) & "] WHERE unit='" + rdr.Item("unit").ToString + "' AND id='" + rdr.Item("id").ToString + "'"
                sqldel.ExecuteNonQuery()

                'See if ID deleted is current ID in units file and if so back off one
                If IsNumeric(id) Then
                    Dim r As Object
                    sql.CommandText = "SELECT ID FROM edm_units WHERE unit='" + unit + "'"
                    r = sql.ExecuteScalar
                    If IsNumeric(r) Then
                        If CInt(id) = CInt(r) Then
                            id = CStr(CInt(id) - 1)
                            sql.CommandText = "UPDATE edm_units SET id='" + id + "' WHERE unit='" + unit + "'"
                            sql.ExecuteNonQuery()
                        End If
                    End If
                End If
            End If
        Else
            MsgBox("The points table is empty.", MsgBoxStyle.Information)
        End If
        points.Close()
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub Menucomport_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Menucomport.Click
        Dim form1 As New frmCommunications
        form1.ShowDialog()
    End Sub

    Private Sub Menuadddatum_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Menuadddatum.Click
        Dim form1 As New frmDatums
        form1.ShowDialog()
    End Sub

    Private Sub Menuinitdirect_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Menuinitdirect.Click
        Dim form1 As New frmInitializeDirect
        form1.ShowDialog()
    End Sub

    Private Sub Menuinitonepoint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Menuinitonepoint.Click
        Dim form1 As New frmonepoint
        form1.ShowDialog()
    End Sub

    Private Sub Menuinitthreepoint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Menuinitthreepoint.Click
        Dim form1 As New frmThreePoint
        form1.ShowDialog()
    End Sub

    Private Sub Menuverify_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Menuverify.Click
        Dim form1 As New frmVerify
        form1.ShowDialog()
    End Sub

    Private Sub Menucurrent_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Menucurrent.Click

        Dim s As String
        s = " X = " & FormatNumber(currentstationx, 3) & Chr(13)
        s = s & " Y = " & FormatNumber(currentstationy, 3) & Chr(13)
        s = s & " Z = " & FormatNumber(currentstationz, 3) & Chr(13)
        s = s & "Name = " & currentstationname
        MsgBox("The station is currently at " & s, MsgBoxStyle.OkOnly, "EDM-Mobile")

    End Sub

    Private Sub Menustatus_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Menustatus.Click

        Dim form1 As New frmStatus
        form1.ShowDialog()

    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuViewEDMINI.Click

        If edmini = "" Then
            MsgBox("Error: An EDMINI file has not yet been initialize.")
            Exit Sub
        End If
        If Not File.Exists(edmini) Then
            MsgBox("Error: '" & edmini & "' does not exist.")
            Exit Sub
        End If
        Dim file1 As New System.IO.StreamReader(edmini)
        Dim form1 As New frmViewer
        form1.TextBox1.Text = file1.ReadToEnd
        file1.Close()
        file1.Dispose()
        form1.Text = "View EDM.INI"
        form1.ShowDialog()

    End Sub

    Private Sub MenuViewCFG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuViewCFG.Click

        If cfgfile = "" Then
            MsgBox("Error: A CFG has not yet been initialize.")
            Exit Sub
        End If
        If Not File.Exists(cfgfile) Then
            MsgBox("Error: '" & cfgfile & "' does not exist.")
            Exit Sub
        End If
        Dim file1 As New System.IO.StreamReader(cfgfile)
        Dim form1 As New frmViewer
        form1.TextBox1.Text = file1.ReadToEnd
        file1.Close()
        file1.Dispose()
        form1.Text = "View CFG File"
        form1.ShowDialog()

    End Sub

    Private Sub recordpoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles recordpoint.Click

        Cursor.Current = Cursors.WaitCursor
        button_no = 0
        shottype = -1
        Dim whichprism As New frmWhichPrism
        If whichprism.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            whichprism.Dispose()
            Dim editpoint As New frmEditPoint
            editpoint.ShowDialog()
            editpoint.Dispose()
        Else
            If commcontrol.IsOpen Then commcontrol.DiscardInBuffer()
        End If
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub Menudatums_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Menudatums.Click
        Dim form1 As New frmDatums
        form1.ShowDialog()
    End Sub

    Private Sub MenuItem2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        Dim form1 As New frmDBGrids
        form1.ShowDialog()

    End Sub

    Private Sub continuepoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles continuepoint.Click

        Cursor.Current = Cursors.WaitCursor
        button_no = 0
        shottype = -8
        Dim whichprism As New frmWhichPrism
        If whichprism.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            whichprism.Dispose()
            Dim editpoint As New frmEditPoint
            editpoint.ShowDialog()
            editpoint.Dispose()
        Else
            If commcontrol.IsOpen Then commcontrol.DiscardInBuffer()
        End If
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub measurepoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles measurepoint.Click

        Cursor.Current = Cursors.WaitCursor
        button_no = 0
        shottype = -2
        Dim whichprism As New frmWhichPrism

        If whichprism.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            currentunit = ""
            If use_limitchecking Then

                Dim d As Single
                Dim units As New SqlCeConnection("datasource='" + options(1) + "'")
                Dim sqlunits As SqlCeCommand = units.CreateCommand
                Dim flag As Boolean = False

                Try
                    units.Open()
                    sqlunits.CommandText = "SELECT * FROM edm_units"
                    Dim rdr As SqlCeDataReader = sqlunits.ExecuteReader
                    Do While rdr.Read
                        If Not IsDBNull(rdr.Item("radius")) Then
                            If CSng(rdr.Item("radius")) <> 0 Then
                                d = CSng(System.Math.Sqrt((CDbl(rdr.Item("centerx")) - currentxp) ^ 2 + (CDbl(rdr.Item("centery")) - currentyp) ^ 2))
                                If d < CSng(rdr.Item("radius")) Then
                                    flag = True
                                    Exit Do
                                End If
                            End If
                        End If
                        If Not IsDBNull(rdr.Item("minx")) And Not IsDBNull(rdr.Item("miny")) And Not IsDBNull(rdr.Item("maxx")) And Not IsDBNull(rdr.Item("maxy")) Then
                            If CDbl(rdr.Item("minx")) <= currentxp And CDbl(rdr.Item("maxx")) >= currentxp And CDbl(rdr.Item("miny")) <= currentyp And CDbl(rdr.Item("maxy")) >= currentyp Then
                                flag = True
                                Exit Do
                            End If
                        End If
                    Loop
                    If flag Then currentunit = rdr.Item("unit").ToString
                    rdr.Close()

                Catch sqlex As System.Data.SqlServerCe.SqlCeException
                    MsgBox("Error reading units table in " & options(1) & vbNewLine & sqlex.Message, MsgBoxStyle.Exclamation, "Issue")

                Finally
                    units.Close()

                End Try

            End If

            Dim info As String = " X = " & FormatNumber(currentxp, 3) & Chr(13)
            info = info & " Y = " & FormatNumber(currentyp, 3) & Chr(13)
            info = info & " Z = " & FormatNumber(currentzp, 3) & Chr(13)
            info = info & " Horizonal angle = " & FormatNumber(currenthangle, 4) & Chr(13)
            info = info & " Vertical angle = " & FormatNumber(currentvangle, 4) & Chr(13)
            info = info & " Slope distance = " & FormatNumber(currentsloped, 3) & Chr(13)
            If currentunit <> "" Then info = info & " Unit = " & currentunit

            Cursor.Current = Cursors.Default
            MsgBox(info, MsgBoxStyle.Information And MsgBoxStyle.OkCancel, "EDM-Mobile")

        Else
            Cursor.Current = Cursors.Default
            If commcontrol.IsOpen Then commcontrol.DiscardInBuffer()

        End If

    End Sub

    Private Sub Menuinittwopoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Menuinittwopoint.Click

        Dim twopoint As New frmtwopoint
        twopoint.ShowDialog()

    End Sub

    Private Sub Menuinitangleonly_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Menuinitangleonly.Click

        Dim angleonly As New frmAngleOnly
        angleonly.ShowDialog()

    End Sub
    Private Sub setup_buttons()

        Dim inidata(100, 2) As String
        Dim iniclass As String
        Dim a As Integer
        Dim b As Integer
        Dim status As Integer
        Dim flag As Boolean

        For a = 1 To 6
            For b = 1 To 100
                button_values(a, b) = ""
            Next b
        Next a

        inidata(1, 1) = "TITLE"
        For a = 1 To vars
            inidata(a + 1, 1) = varlist(a)
        Next a

        flag = False
        For a = 1 To 6
            iniclass = "[BUTTON" & Trim(CStr(a)) & "]"
            For b = 1 To vars + 1
                inidata(b, 2) = ""
            Next b
            Call ReadIni(cfgfile, iniclass, inidata, status)
            If inidata(1, 2) <> "" Then
                Select Case a
                    Case 1
                        Button1.Text = inidata(1, 2)
                        Button1.Visible = True
                        flag = True
                    Case 2
                        Button2.Text = inidata(1, 2)
                        Button2.Visible = True
                        flag = True
                    Case 3
                        Button3.Text = inidata(1, 2)
                        Button3.Visible = True
                        flag = True
                    Case 4
                        Button4.Text = inidata(1, 2)
                        Button4.Visible = True
                        flag = True
                    Case 5
                        Button5.Text = inidata(1, 2)
                        Button5.Visible = True
                        flag = True
                    Case 6
                        Button6.Text = inidata(1, 2)
                        Button6.Visible = True
                        flag = True
                    Case Else
                End Select
                For b = 2 To vars + 1
                    button_values(a, b - 1) = inidata(b, 2)
                Next b
            Else
                Select Case a
                    Case 1
                        Button1.Visible = False
                    Case 2
                        Button2.Visible = False
                    Case 3
                        Button3.Visible = False
                    Case 4
                        Button4.Visible = False
                    Case 5
                        Button5.Visible = False
                    Case 6
                        Button6.Visible = False
                    Case Else
                End Select
            End If
        Next a

        If flag = False Then
            frame2.Visible = False
        Else
            frame2.Visible = True
        End If

    End Sub


    Private Sub Menueditlast_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Menueditlast.Click

        Dim a As Integer
        Dim points As New SqlCeConnection("datasource='" + options(1) + "'")
        Dim sql As SqlCeCommand = points.CreateCommand
        Cursor.Current = Cursors.WaitCursor
        sql.CommandText = "SELECT * FROM [" & options(22) & "] ORDER BY Recno Desc"
        points.Open()
        Dim rdr As SqlCeDataReader = sql.ExecuteReader
        If rdr.Read Then

            'Note the record number
            Dim recno As Integer = CInt(rdr.Item("recno"))

            'Get the values of the last record
            For a = 1 To vars
                responses(a) = rdr.Item(varlist(a)).ToString
            Next
            Dim editpoint As New frmEditPoint
            back_to_main = True
            If editpoint.ShowDialog() = Windows.Forms.DialogResult.OK Then

                'Fourth, put these new responses back into the record
                Dim id As String = ""
                Dim unit As String = ""
                sql.CommandText = "UPDATE [" & options(22) & "] SET "
                For a = 1 To vars
                    Select Case vtype(a)
                        Case 2, 3
                            If a <> vars Then
                                If responses(a) <> "" Then
                                    sql.CommandText = sql.CommandText + varlist(a) + "=" + responses(a) + ","
                                Else
                                    sql.CommandText = sql.CommandText + varlist(a) + "=0,"
                                End If
                            Else
                                If responses(a) <> "" Then
                                    sql.CommandText = sql.CommandText + varlist(a) + "=" + responses(a)
                                Else
                                    sql.CommandText = sql.CommandText + varlist(a) + "=0"
                                End If
                            End If
                        Case Else
                            If a <> vars Then
                                sql.CommandText = sql.CommandText + varlist(a) + "='" + responses(a) + "',"
                            Else
                                sql.CommandText = sql.CommandText + varlist(a) + "='" + responses(a) + "'"
                            End If
                    End Select
                    If UCase(varlist(a)) = "UNIT" Then unit = responses(a)
                    If UCase(varlist(a)) = "ID" Then id = responses(a)
                Next a
                sql.CommandText = sql.CommandText + " WHERE recno=" & recno
                sql.ExecuteNonQuery()

                'Fifth, make sure Unit information has the last ID
                If IsNumeric(id) Then
                    Dim r As Object
                    sql.CommandText = "SELECT ID FROM edm_units WHERE unit='" + unit + "'"
                    r = sql.ExecuteScalar
                    If IsNumeric(r) Then
                        If CInt(id) > CInt(r) Then
                            sql.CommandText = "UPDATE edm_units SET id='" + id + "' WHERE unit='" + unit + "'"
                            sql.ExecuteNonQuery()
                        End If
                    End If
                End If
            End If
        Else
            MsgBox("The points table is empty.", MsgBoxStyle.Information)
        End If
        points.Close()
        back_to_main = False
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub MenuItem23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem23.Click

        Select Case options(4)
            Case "TOPCON", "WILD", "SOKKIA", "LEICA", "NIKON", "LEICA-GEO"
                Cursor.Current = Cursors.WaitCursor
                Call delay(10)
                Cursor.Current = Cursors.Default
                Call close_com()
                Call open_com()
                MsgBox("The connection is reset.  Use this option when the computer powers off while the program is running.", vbOKOnly, "EDM-Mobile")
            Case Else
                MsgBox("A valid total station type has not yet been selected so the connection can not be reset.", vbOKOnly, "EDM-Mobile")
        End Select

    End Sub

    Private Sub Menusleep_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Menusleep.Click

        If MsgBox("Put computer into stand-by mode?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "EDM-Mobile") = MsgBoxResult.No Then Exit Sub

        Call close_com()
        SetSystemPowerState(Nothing, POWER_STATE_SUSPEND, Nothing)
        'SetSystemPowerState(Nothing, POWER_STATE_SUSPEND, POWER_FORCE)
        Select Case options(4)
            Case "TOPCON", "WILD", "SOKKIA", "LEICA", "NIKON", "LEICA-GEO"
                MsgBox("Press Ok and the program will reset the communication port.  Please allow up to 30 seconds (even if it appears to have crashed).", MsgBoxStyle.OkOnly, "EDM-Mobile")
                Cursor.Current = Cursors.WaitCursor
                Call delay(10)
                Call open_com()
                Cursor.Current = Cursors.Default
            Case Else
        End Select

    End Sub

    Private Sub Menuexport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Menuexport.Click

        If MsgBox("This option exports the points, datums, units, and prism tables to ASCII files.  When asked for a file name, enter a root name and '_points', '_datums', etc. will added to the root.", MsgBoxStyle.OkCancel) = MsgBoxResult.Cancel Then Exit Sub

        sfd.Filter = "Text File|*.txt"
        sfd.ShowDialog()

        If sfd.FileName <> "" Then

            Dim r As String = sfd.FileName
            Dim a As Integer = InStr(r, ".")
            If a <> 0 Then r = leftstr(r, a - 1)
            Dim f As String = r + "_points.txt"
            Dim write_header As MsgBoxResult

            If File.Exists(f) Then
                If MsgBox("Overwrite existing files?", vbOKCancel, "EDMCE") = MsgBoxResult.Cancel Then Exit Sub
            End If

            write_header = MsgBox("Include a header record?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question)

            Cursor.Current = Cursors.WaitCursor

            Dim db As New SqlCeConnection("datasource='" + options(1) + "'")
            Dim sql As SqlCeCommand = db.CreateCommand
            Dim rdr As SqlCeDataReader

            db.Open()
            sql.CommandText = "SELECT * FROM [" & options(22) & "] ORDER BY Recno"
            rdr = sql.ExecuteReader
            Call export_ascii(rdr, f, write_header)
            rdr.Close()

            Dim fd As String = r + "_datums.txt"
            sql.CommandText = "SELECT * FROM edm_datums ORDER BY name"
            rdr = sql.ExecuteReader
            Call export_ascii(rdr, fd, write_header)
            rdr.Close()

            Dim fp As String = r + "_prisms.txt"
            sql.CommandText = "SELECT * FROM edm_poles ORDER BY name"
            rdr = sql.ExecuteReader
            Call export_ascii(rdr, fp, write_header)
            rdr.Close()

            Dim fu As String = r + "_units.txt"
            sql.CommandText = "SELECT * FROM edm_units ORDER BY unit"
            rdr = sql.ExecuteReader
            Call export_ascii(rdr, fu, write_header)
            rdr.Close()
            db.Close()

            Cursor.Current = Cursors.Default

            MsgBox("ASCII files " + f + "," + fp + "," + fd + "," + fu + " were written.", MsgBoxStyle.Information And MsgBoxStyle.OkOnly)

        End If

    End Sub

    Private Sub menuspeedbuttons_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles menuspeedbuttons.Click

        Dim speedbuttons As New frmButtons
        speedbuttons.ShowDialog()
        Call setup_buttons()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Cursor.Current = Cursors.WaitCursor
        button_no = 1
        shottype = -1
        Dim whichprism As New frmWhichPrism
        If whichprism.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            whichprism.Dispose()
            Dim editpoint As New frmEditPoint
            editpoint.ShowDialog()
            editpoint.Dispose()
        Else
            If commcontrol.IsOpen Then commcontrol.DiscardInBuffer()
        End If
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Cursor.Current = Cursors.WaitCursor
        button_no = 2
        shottype = -1
        Dim whichprism As New frmWhichPrism
        If whichprism.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            whichprism.Dispose()
            Dim editpoint As New frmEditPoint
            editpoint.ShowDialog()
            editpoint.Dispose()
        Else
            If commcontrol.IsOpen Then commcontrol.DiscardInBuffer()
        End If
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Cursor.Current = Cursors.WaitCursor
        button_no = 3
        shottype = -1
        Dim whichprism As New frmWhichPrism
        If whichprism.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            whichprism.Dispose()
            Dim editpoint As New frmEditPoint
            editpoint.ShowDialog()
            editpoint.Dispose()
        Else
            If commcontrol.IsOpen Then commcontrol.DiscardInBuffer()
        End If
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Cursor.Current = Cursors.WaitCursor
        button_no = 4
        shottype = -1
        Dim whichprism As New frmWhichPrism
        If whichprism.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            whichprism.Dispose()
            Dim editpoint As New frmEditPoint
            editpoint.ShowDialog()
            editpoint.Dispose()
        Else
            If commcontrol.IsOpen Then commcontrol.DiscardInBuffer()
        End If
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        Cursor.Current = Cursors.WaitCursor
        button_no = 5
        shottype = -1
        Dim whichprism As New frmWhichPrism
        If whichprism.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            whichprism.Dispose()
            Dim editpoint As New frmEditPoint
            editpoint.ShowDialog()
            editpoint.Dispose()
        Else
            If commcontrol.IsOpen Then commcontrol.DiscardInBuffer()
        End If
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click

        Cursor.Current = Cursors.WaitCursor
        button_no = 6
        shottype = -1
        Dim whichprism As New frmWhichPrism
        If whichprism.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            whichprism.Dispose()
            Dim editpoint As New frmEditPoint
            editpoint.ShowDialog()
            editpoint.Dispose()
        Else
            If commcontrol.IsOpen Then commcontrol.DiscardInBuffer()
        End If
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub Menuabout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Menuabout.Click

        Dim about As New frmAbout
        about.ShowDialog()

    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click

        If cfgfile = "" Then
            MsgBox("First open a configuration file before making a Log file.", MsgBoxStyle.Information And MsgBoxStyle.OkOnly)
            Exit Sub
        End If

        If Not WriteLogFile Then
            If MsgBox("The option to keep a Log file is currently not activated (see Edit-EDM-Mobile Options).  Would you like to activate it now?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                WriteLogFile = True
            End If
        End If

        If LogFile = "" Then
            LogFile = Path.GetDirectoryName(cfgfile) + "\" + Path.GetFileNameWithoutExtension(cfgfile) + "_log.txt"
        End If

        If MsgBox("Erase the contents of the Log file " + LogFile + "?", MsgBoxStyle.OkCancel) = MsgBoxResult.Cancel Then Exit Sub

        If File.Exists(LogFile) Then File.Delete(LogFile)

        Dim file3 As StreamWriter = File.CreateText(LogFile)
        file3.WriteLine("Log File")
        file3.Flush()
        file3.Close()
        file3.Dispose()

    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click

        If LogFile = "" Then
            MsgBox("Error: A Log file has not yet been created.  Open a CFG first.")
            Exit Sub
        End If

        If Not WriteLogFile Then
            If MsgBox("The option to keep a Log file is currently not activated (see Edit-EDM-Mobile Options).  Would you like to activate it now?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                WriteLogFile = True
                Dim iniclass As String
                Dim inidata(3, 2) As String
                iniclass = "[EDM]"
                inidata(1, 1) = "WriteLogFile"
                inidata(3, 2) = "Yes"
                Dim status As Integer
                Call WriteIni(cfgfile, iniclass, inidata, status)
            Else
                Exit Sub
            End If
        End If

        If Not File.Exists(LogFile) Then
            MsgBox("Error: '" & LogFile & "' does not yet exist.")
            Exit Sub
        End If

        Dim file1 As New System.IO.StreamReader(LogFile)
        Dim form1 As New frmViewer
        form1.TextBox1.Text = file1.ReadToEnd
        file1.Close()
        file1.Dispose()
        form1.Text = "View Log File"
        form1.ShowDialog()

    End Sub

    Private Sub unit_id_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles unit_id.TextChanged

        bar_code_unit_id = unit_id.Text

    End Sub

    Private Sub Battery_test_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Battery_test.Click

        Select Case options(4)
            Case "TOPCON", "WILD", "SOKKIA", "LEICA", "NIKON", "SIMULATION", "LEICA-GEO"
                If options(4) = "SIMULATION" Then MsgBox("Warning: You are in simulation mode.  No actual shots will be taken.  Use Station-Settings to switch total station type.", MsgBoxStyle.Information)
                Dim form1 As New frmBatteryTest
                form1.ShowDialog()
            Case Else
                MsgBox("A valid total station type has not yet been selected. Select one to test the battery on the station.", vbOKOnly, "EDM-Mobile")
        End Select

    End Sub

    Private Sub MenuViewPoints_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuViewPoints.Click

        Cursor.Current = Cursors.WaitCursor
        Dim form1 As New frmPlots
        form1.ShowDialog()
        Cursor.Current = Cursors.Default

    End Sub
End Class
