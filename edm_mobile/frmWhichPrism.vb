Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlServerCe
Imports System.io
Imports System
Public Class frmWhichPrism
    Inherits System.Windows.Forms.Form
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents prismlist As System.Windows.Forms.ListBox
    Friend WithEvents hangle As System.Windows.Forms.Label
    Friend WithEvents vangle As System.Windows.Forms.Label
    Friend WithEvents sloped As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents prismconstantlabel As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents zresult As System.Windows.Forms.Label
    Friend WithEvents yresult As System.Windows.Forms.Label
    Friend WithEvents xresult As System.Windows.Forms.Label
    Friend WithEvents dontwait As System.Windows.Forms.CheckBox
    Friend WithEvents barcodetrap As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button

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
    Friend WithEvents newprismheight As System.Windows.Forms.TextBox
    Friend WithEvents newprismconstant As System.Windows.Forms.TextBox
    Friend WithEvents newprismname As System.Windows.Forms.TextBox
    Friend WithEvents reflector As System.Windows.Forms.Label
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.prismlist = New System.Windows.Forms.ListBox
        Me.hangle = New System.Windows.Forms.Label
        Me.vangle = New System.Windows.Forms.Label
        Me.sloped = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.newprismheight = New System.Windows.Forms.TextBox
        Me.newprismconstant = New System.Windows.Forms.TextBox
        Me.prismconstantlabel = New System.Windows.Forms.Label
        Me.newprismname = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.reflector = New System.Windows.Forms.Label
        Me.Timer1 = New System.Windows.Forms.Timer
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.zresult = New System.Windows.Forms.Label
        Me.yresult = New System.Windows.Forms.Label
        Me.xresult = New System.Windows.Forms.Label
        Me.dontwait = New System.Windows.Forms.CheckBox
        Me.barcodetrap = New System.Windows.Forms.TextBox
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(216, 16)
        Me.Label1.Text = "Select a prism or enter a new one:"
        '
        'prismlist
        '
        Me.prismlist.Location = New System.Drawing.Point(8, 32)
        Me.prismlist.Name = "prismlist"
        Me.prismlist.Size = New System.Drawing.Size(92, 100)
        Me.prismlist.TabIndex = 12
        '
        'hangle
        '
        Me.hangle.Location = New System.Drawing.Point(8, 138)
        Me.hangle.Name = "hangle"
        Me.hangle.Size = New System.Drawing.Size(112, 16)
        Me.hangle.Text = "H Angle 180.0000"
        '
        'vangle
        '
        Me.vangle.Location = New System.Drawing.Point(8, 155)
        Me.vangle.Name = "vangle"
        Me.vangle.Size = New System.Drawing.Size(112, 16)
        Me.vangle.Text = "V Angle 180.0000"
        '
        'sloped
        '
        Me.sloped.Location = New System.Drawing.Point(8, 171)
        Me.sloped.Name = "sloped"
        Me.sloped.Size = New System.Drawing.Size(112, 16)
        Me.sloped.Text = "Slope D. 9999.999"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(117, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(41, 16)
        Me.Label2.Text = "Height"
        '
        'newprismheight
        '
        Me.newprismheight.Location = New System.Drawing.Point(160, 56)
        Me.newprismheight.MaxLength = 15
        Me.newprismheight.Name = "newprismheight"
        Me.newprismheight.Size = New System.Drawing.Size(59, 21)
        Me.newprismheight.TabIndex = 6
        Me.newprismheight.Text = "0.000"
        '
        'newprismconstant
        '
        Me.newprismconstant.Location = New System.Drawing.Point(160, 32)
        Me.newprismconstant.MaxLength = 15
        Me.newprismconstant.Name = "newprismconstant"
        Me.newprismconstant.Size = New System.Drawing.Size(43, 21)
        Me.newprismconstant.TabIndex = 4
        Me.newprismconstant.Text = "0.000"
        '
        'prismconstantlabel
        '
        Me.prismconstantlabel.Location = New System.Drawing.Point(105, 32)
        Me.prismconstantlabel.Name = "prismconstantlabel"
        Me.prismconstantlabel.Size = New System.Drawing.Size(53, 13)
        Me.prismconstantlabel.Text = "Constant"
        '
        'newprismname
        '
        Me.newprismname.Location = New System.Drawing.Point(160, 80)
        Me.newprismname.MaxLength = 20
        Me.newprismname.Name = "newprismname"
        Me.newprismname.Size = New System.Drawing.Size(72, 21)
        Me.newprismname.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(119, 80)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(39, 21)
        Me.Label4.Text = "Name"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(6, 219)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(229, 40)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Continue"
        '
        'reflector
        '
        Me.reflector.ForeColor = System.Drawing.Color.Red
        Me.reflector.Location = New System.Drawing.Point(38, 192)
        Me.reflector.Name = "reflector"
        Me.reflector.Size = New System.Drawing.Size(162, 16)
        Me.reflector.Text = "Reflectorless Mode"
        Me.reflector.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Timer1
        '
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.Add(Me.MenuItem1)
        Me.MainMenu1.MenuItems.Add(Me.MenuItem3)
        '
        'MenuItem1
        '
        Me.MenuItem1.Text = "Reshoot"
        '
        'MenuItem3
        '
        Me.MenuItem3.Text = "Cancel"
        '
        'zresult
        '
        Me.zresult.Location = New System.Drawing.Point(130, 171)
        Me.zresult.Name = "zresult"
        Me.zresult.Size = New System.Drawing.Size(103, 16)
        Me.zresult.Text = "Z : 99999999.999"
        Me.zresult.Visible = False
        '
        'yresult
        '
        Me.yresult.Location = New System.Drawing.Point(130, 155)
        Me.yresult.Name = "yresult"
        Me.yresult.Size = New System.Drawing.Size(103, 16)
        Me.yresult.Text = "Y : 99999999.999"
        Me.yresult.Visible = False
        '
        'xresult
        '
        Me.xresult.Location = New System.Drawing.Point(130, 138)
        Me.xresult.Name = "xresult"
        Me.xresult.Size = New System.Drawing.Size(103, 16)
        Me.xresult.Text = "X : 99999999.999"
        Me.xresult.Visible = False
        '
        'dontwait
        '
        Me.dontwait.Location = New System.Drawing.Point(106, 110)
        Me.dontwait.Name = "dontwait"
        Me.dontwait.Size = New System.Drawing.Size(126, 21)
        Me.dontwait.TabIndex = 20
        Me.dontwait.Text = "Auto select last"
        '
        'barcodetrap
        '
        Me.barcodetrap.Location = New System.Drawing.Point(221, 32)
        Me.barcodetrap.Name = "barcodetrap"
        Me.barcodetrap.Size = New System.Drawing.Size(11, 21)
        Me.barcodetrap.TabIndex = 32
        Me.barcodetrap.Visible = False
        '
        'frmWhichPrism
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(240, 270)
        Me.ControlBox = False
        Me.Controls.Add(Me.barcodetrap)
        Me.Controls.Add(Me.dontwait)
        Me.Controls.Add(Me.zresult)
        Me.Controls.Add(Me.yresult)
        Me.Controls.Add(Me.xresult)
        Me.Controls.Add(Me.reflector)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.newprismname)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.newprismconstant)
        Me.Controls.Add(Me.prismconstantlabel)
        Me.Controls.Add(Me.newprismheight)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.sloped)
        Me.Controls.Add(Me.vangle)
        Me.Controls.Add(Me.hangle)
        Me.Controls.Add(Me.prismlist)
        Me.Controls.Add(Me.Label1)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmWhichPrism"
        Me.Text = "Prisms"
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim edmdata As String

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim a As Integer
        Dim flag As Boolean
        Dim p As String

        Cursor.Current = Cursors.WaitCursor

        If newprismname.Text <> "" Then

            npoles = 0  'force a re-read of temporary array storing pole data

            If Not IsNumeric(newprismheight.Text) Then
                Cursor.Current = Cursors.Default
                MsgBox("To add a prism you must supply a height.", MsgBoxStyle.Information, "EDM-Mobile")
                Exit Sub
            End If

            If Not IsNumeric(newprismconstant.Text) Then
                currentprismoffset = 0
            Else
                currentprismoffset = CSng(newprismconstant.Text)
            End If

            If currentprismoffset > 0.2 Or currentprismoffset < -0.2 Then
                MsgBox("Warning: Prism offset values are typically less than 20cm. Make sure offset and height have not been confused.", MsgBoxStyle.Information And MsgBoxStyle.OkOnly)
            End If

            currentprismname = newprismname.Text
            currentprismheight = CSng(newprismheight.Text)

            flag = False
            For a = 0 To prismlist.Items.Count - 1
                If LCase(CStr(prismlist.Items.Item(a))) = LCase(newprismname.Text) Then
                    flag = True
                End If
            Next a

            'check to see if this prism exists already
            p = newprismname.Text
            flag = False
            Dim prisms As New SqlCeConnection("datasource =" + options(1))
            Dim sqlprisms As SqlCeCommand = prisms.CreateCommand
            sqlprisms.CommandText = "SELECT name FROM edm_poles WHERE name='" + p + "'"
            prisms.Open()

            'delete it if it does
            Dim r As Object = sqlprisms.ExecuteScalar()
            If LCase(CStr(r)) = LCase(p) Then
                sqlprisms.CommandText = "DELETE FROM edm_poles WHERE name='" + p + "'"
                sqlprisms.ExecuteNonQuery()
            End If

            'update and/or add the new one
            sqlprisms.CommandText = "INSERT INTO edm_poles (name,height,offset) VALUES ('" + p + "'," + CStr(currentprismheight) + "," + CStr(currentprismoffset) + ")"
            sqlprisms.ExecuteNonQuery()
            sqlprisms.Dispose()
            prisms.Close()
            prisms.Dispose()

        ElseIf prismlist.SelectedIndex > -1 Then
            currentprismname = prismlist.SelectedItem.ToString
            If newprismconstant.Text = "" Then
                currentprismoffset = 0
            Else
                currentprismoffset = CSng(newprismconstant.Text)
            End If
            currentprismheight = CSng(newprismheight.Text)

        ElseIf IsNumeric(newprismheight.Text) Then
            currentprismheight = CSng(newprismheight.Text)
            If IsNumeric(newprismconstant.Text) Then
                currentprismoffset = CSng(newprismconstant.Text)
            Else
                currentprismoffset = 0
            End If
            currentprismname = ""

        Else
            currentprismoffset = 0
            currentprismname = "Zero"
            currentprismheight = 0

        End If

        'Now that the prism offset is known - calculate the XYZ from the sloped, h and v angles
        If options$(4) <> "SIMULATION" Then Call vhdtonez()

        shot_canceled = False

        currentzp = CDbl(FormatNumber(currentzp - currentprismheight, 3))
        currentxp = CDbl(FormatNumber(currentxp, 3))
        currentyp = CDbl(FormatNumber(currentyp, 3))

        If shottype = -1 Then
            'real shot
            currentxp = CDbl(currentstationx) + CDbl(currentxp)
            currentyp = CDbl(currentstationy) + CDbl(currentyp)
            currentzp = CDbl(currentstationz) + CDbl(currentzp)
            Me.DialogResult = Windows.Forms.DialogResult.OK

        ElseIf shottype = 1 Then 'coming from edit form
            currentxp = CDbl(currentstationx) + CDbl(currentxp)
            currentyp = CDbl(currentstationy) + CDbl(currentyp)
            currentzp = CDbl(currentstationz) + CDbl(currentzp)
            Me.DialogResult = Windows.Forms.DialogResult.OK

        ElseIf shottype = -2 Then 'x-shot
            currentxp = CDbl(currentstationx) + CDbl(currentxp)
            currentyp = CDbl(currentstationy) + CDbl(currentyp)
            currentzp = CDbl(currentstationz) + CDbl(currentzp)
            Me.DialogResult = Windows.Forms.DialogResult.OK

        ElseIf shottype = -3 Then 'one point setup shot
            Cursor.Current = Cursors.Default
            Me.DialogResult = Windows.Forms.DialogResult.OK

        ElseIf shottype = -4 Then 'two point setup shot 
            Cursor.Current = Cursors.Default
            Me.DialogResult = Windows.Forms.DialogResult.OK

        ElseIf shottype = -5 Then ' three point setup
            Me.DialogResult = Windows.Forms.DialogResult.OK

        ElseIf shottype = -6 Then
            currentxp = CDbl(currentstationx) + CDbl(currentxp)
            currentyp = CDbl(currentstationy) + CDbl(currentyp)
            currentzp = CDbl(currentstationz) + CDbl(currentzp)
            Cursor.Current = Cursors.Default
            Me.DialogResult = DialogResult.OK

        ElseIf shottype = -7 Then
            currentxp = CDbl(currentstationx) + CDbl(currentxp)
            currentyp = CDbl(currentstationy) + CDbl(currentyp)
            currentzp = CDbl(currentstationz) + CDbl(currentzp)
            Dim pointsform As New frmPoints
            pointsform.Show()

        ElseIf shottype = -8 Then 'continuation shot
            currentxp = CDbl(currentstationx) + CDbl(currentxp)
            currentyp = CDbl(currentstationy) + CDbl(currentyp)
            currentzp = CDbl(currentstationz) + CDbl(currentzp)
            Me.DialogResult = Windows.Forms.DialogResult.OK

        ElseIf shottype = -9 Then 'station verify
            currentxp = CDbl(currentstationx) + CDbl(currentxp)
            currentyp = CDbl(currentstationy) + CDbl(currentyp)
            currentzp = CDbl(currentstationz) + CDbl(currentzp)
            Me.DialogResult = Windows.Forms.DialogResult.OK

        End If

    End Sub

    Private Sub frmWhichPrism_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim angle As Double
        Dim a As Integer

        xresult.Visible = False
        yresult.Visible = False
        zresult.Visible = False
        hangle.Visible = False
        vangle.Visible = False
        sloped.Visible = False
        'prismc.Visible = False
        reflector.Visible = False
        If showoffsetmenu Then
            prismconstantlabel.Visible = True
            newprismconstant.Visible = True
        Else
            prismconstantlabel.Visible = False
            newprismconstant.Visible = False
        End If
        Button1.Enabled = False

        Me.Show()

        shot_canceled = False

        'start the recording process
        edmdata = ""
        Select Case options(4)
            Case "TOPCON", "WILD", "SOKKIA", "LEICA", "LEICA-GEO"
                Call take_shot()
                Timer1.Enabled = True

            Case "SIMULATION"

                'rigged to fall into 2 squares
                If Rnd() > 0.5 Then
                    currentxp = Rnd()
                    currentyp = Rnd()
                    currentzp = Rnd()
                Else
                    currentxp = Rnd() + 1
                    currentyp = Rnd() + 1
                    currentzp = Rnd() + 1
                End If

                Call calculate_angle(currentstationx, currentstationy, currentxp, currentyp, angle)
                currenthangle = CStr(angle)
                currentsloped = System.Math.Sqrt((currentstationx - currentxp) ^ 2 + (currentstationy - currentyp) ^ 2 + (currentstationz - currentzp) ^ 2)
                Call calculate_angle(0, 0, currentsloped, currentzp, angle)
                currentvangle = CStr(angle)
                Button1.Enabled = True
                Timer1.Enabled = False

            Case Else
                MsgBox("The total station types must be set to either Topcon, Wild, Sokkia, Leica, LeicaGeo or Simulation in order to record a point.  Check Station-Total Station Type", MsgBoxStyle.Information And MsgBoxStyle.OkOnly, "EDM-Mobile")
        End Select

        prismlist.BeginUpdate()
        prismlist.Items.Clear()

        If options(1) = "" Then
            MsgBox("Error creating prism table.  Empty database name.", MsgBoxStyle.Exclamation, "Issue")
        Else
            If npoles = 0 Then
                Dim prisms As New SqlCeConnection("datasource='" + options(1) + "'")
                Dim sql As SqlCeCommand = prisms.CreateCommand

                newprismheight.Text = "0.00"
                newprismconstant.Text = "0.00"
                newprismname.Text = ""

                Try
                    prisms.Open()
                    sql.CommandText = "SELECT * FROM edm_poles"
                    Dim rdr As SqlCeDataReader = sql.ExecuteReader
                    prismlist.Items.Clear()
                    npoles = 0
                    Do While rdr.Read
                        prismlist.Items.Add(rdr.Item("name"))
                        npoles = npoles + 1
                        polename(npoles) = rdr.Item("name").ToString
                        poledata(npoles, 1) = CSng(rdr.Item("height"))
                        poledata(npoles, 2) = CSng(rdr.Item("offset"))
                    Loop
                    rdr.Close()

                Catch sqlex As System.Data.SqlServerCe.SqlCeException
                    MsgBox("Error reading prisms table in " & options(1) & vbNewLine & sqlex.Message, MsgBoxStyle.Exclamation, "Issue")
                    Exit Try

                Finally
                    prisms.Close()

                End Try
            Else
                For a = 1 To npoles
                    prismlist.Items.Add(polename(a))
                Next
            End If
        End If

        prismlist.EndUpdate()

        For a = 0 To prismlist.Items.Count - 1
            If prismlist.Items(a).ToString = currentprismname Then
                prismlist.SelectedIndex = a
                Exit For
            End If
        Next a

        If donotwaitforprism Then dontwait.Checked = True

        barcodetrap.Focus()

        Cursor.Current = Cursors.Default

    End Sub

    Private Sub prismlist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles prismlist.SelectedIndexChanged

        If prismlist.SelectedIndex <> -1 Then
            Dim p As String
            Dim prisms As New SqlCeConnection("datasource='" + options(1) + "'")
            Dim sql As SqlCeCommand = prisms.CreateCommand
            Cursor.Current = Cursors.WaitCursor
            p = prismlist.SelectedItem.ToString
            sql.CommandText = "SELECT name,height,offset FROM edm_poles WHERE name='" + p + "'"
            prisms.Open()
            Dim rdr As SqlCeDataReader = sql.ExecuteReader
            Do While rdr.Read
                newprismheight.Text = Format(rdr.Item("height"), "##0.000")
                newprismconstant.Text = Format(rdr.Item("offset"), "0.000")
                currentprismname = rdr.Item("name").ToString
            Loop
            prisms.Close()
            If zresult.Visible Then zresult.Text = "Z: " & CDbl(currentstationz) + CDbl(FormatNumber(currentzp - CSng(newprismheight.Text), 3))
            Cursor.Current = Cursors.Default
        End If

    End Sub

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Dim ack As String = ""
        Dim returndata As String = ""
        Dim edmpoffset As Single
        Dim mesunits As String = ""
        Dim angleunit As String = ""
        Dim t As String
        Dim status As Integer
        Dim errorcode As Integer

        If options$(4) = "SIMULATION" Then

            edmdata = "R+00001432m0960005+3225010d-00000151t34+00-01014" + Chr(13) + Chr(10)

        Else
            'look for incoming data
            't = commcontrol.ReadLine()
            If commcontrol.IsOpen Then
                If commcontrol.BytesToRead > 0 Then
                    t = ""
                    Do
                        t = Chr(commcontrol.ReadChar)
                        If t <> "" Then edmdata = edmdata + t
                    Loop Until commcontrol.BytesToRead = 0 Or t = Chr(10)
                End If
            End If
        End If

        'look for the carriage return at the end of a line
        If Microsoft.VisualBasic.Right(edmdata, 2) = Chr(13) + Chr(10) Then

            Timer1.Enabled = False

            Select Case UCase(options(4))
                Case "TOPCON"
                    ack = Chr(6) + "006" + Chr(3) + Chr(13) + Chr(10)
                    Call directoutput(ack)
                    Do Until Asc(Microsoft.VisualBasic.Left(edmdata, 1)) > 32 Or Len(edmdata) = 1
                        edmdata = Mid(edmdata, 2)
                    Loop
                    returndata = edmdata

                    Call delay(1)

                    Call horizontal(errorcode)
                    Call clearcom()

                Case "WILD", "LEICA", "LEICA-GEO"
                    returndata = edmdata
                    Call clearcom()

                Case "SOKKIA"
                    returndata = edmdata
                    If edmdata = "CANCEL" Then
                        t = Chr(18)
                        Call edmoutput(t, errorcode)
                    End If
                    Call clearcom()

                Case Else
            End Select

            Call parsenez(returndata, edmpoffset, mesunits, angleunit, status)
            edmprismoffset = edmpoffset

            If options(4) = "SIMULATION" Then
                If Rnd() > 0.5 Then
                    currentxp = Rnd()
                    currentyp = Rnd()
                    currentzp = Rnd()
                Else
                    currentxp = Rnd() + 10
                    currentyp = Rnd() + 10
                    currentzp = Rnd() + 10
                End If
            End If

            If status = 0 Then
                Button1.Enabled = True
                hangle.Visible = True
                hangle.Text = "HAngle: " + currenthangle
                vangle.Visible = True
                vangle.Text = "VAngle: " + currentvangle
                sloped.Visible = True
                sloped.Text = "SlopeD: " + FormatNumber(currentsloped, 3)
                'prismc.Visible = True
                'prismc.Text = "Constant:" + CStr(edmpoffset)
                If edmprismoffset = 0.004 And (options(4) = "WILD" Or options(4) = "LEICA" Or options(4) = "LEICA-GEO") Then
                    reflector.Visible = True
                End If
                xresult.Text = "X: " & CDbl(currentstationx) + CDbl(currentxp)
                yresult.Text = "Y: " & CDbl(currentstationy) + CDbl(currentyp)
                zresult.Text = "Z: " & CDbl(currentstationz) + CDbl(FormatNumber(currentzp - currentprismheight, 3))
                'xresult.Visible = True
                'yresult.Visible = True
                'zresult.Visible = True
                If donotwaitforprism Then Call Button1_Click(Timer1, e)
            Else
                Select Case status
                    Case 1, 5
                        Select Case options(4)
                            Case "WILD", "LEICA", "LEICA-GEO"
                                MsgBox("ERROR: Shot not recorded.  Check that the instrument is aimed at the prism and that nothing blocks the laser's path.", vbCritical, "EDM-Mobile")
                            Case Else
                                MsgBox("ERROR: Instrument was not in Vertical distance mode.  Reset instrument.", vbCritical, "EDM-Mobile")
                        End Select
                    Case 6
                        MsgBox("ERROR: Parity error.  Shot must be retaken.", vbCritical, "EDM-Mobile")
                    Case 2, 20
                        MsgBox("ERROR: Could not parse data from station. Check that horizontal angles are in degrees, minutes and seconds.", vbCritical, "EDM-Mobile")
                    Case 100
                        MsgBox("ERROR: The slope distance is 0.  Shot must be retaken.  Usually this means that something is blocking the path to the prism, you are shooting something too far away, or the surface is not reflective enough (in reflectorless mode).", vbCritical, "EDM-Mobile")
                    Case 200
                        MsgBox("ERROR: Hangle set to increase to the left.  Reset to increase to the right.", vbCritical, "EDM-Mobile")
                    Case Else
                        MsgBox("ERROR: Unknown error.  Shot must be retaken or station not configured correctly.", vbCritical, "EDM-Mobile")
                End Select
            End If
            edmdata = ""

        End If

    End Sub

    Private Sub Timer1_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Disposed

    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click

        Dim errorcode As Integer

        Timer1.Enabled = False
        Cursor.Current = Cursors.WaitCursor
        Call horizontal(errorcode)
        Cursor.Current = Cursors.Default
        shot_canceled = True
        Me.DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click

        Dim angle As Double
        hangle.Visible = False
        vangle.Visible = False
        sloped.Visible = False
        'prismc.Visible = False
        reflector.Visible = False
        Button1.Enabled = False
        xresult.Visible = False
        yresult.Visible = False
        zresult.Visible = False

        'start the recording process
        edmdata = ""
        Select Case options(4)
            Case "TOPCON", "WILD", "SOKKIA", "LEICA", "LEICA-GEO"
                Call take_shot()
                Timer1.Enabled = True

            Case "SIMULATION"

                'rigged to fall into 2 squares
                If Rnd() > 0.5 Then 'c12
                    currentxp = Rnd() + 10
                    currentyp = Rnd() + 10
                    currentzp = 100 + Rnd() * 10 - 5
                Else 'e18
                    currentxp = Rnd() + 12
                    currentyp = Rnd() + 12
                    currentzp = 100 + Rnd() * 10 - 5
                End If

                Call calculate_angle(currentstationx, currentstationy, currentxp, currentyp, angle)
                currenthangle = CStr(angle)
                currentsloped = System.Math.Sqrt((currentstationx - currentxp) ^ 2 + (currentstationy - currentyp) ^ 2 + (currentstationz - currentzp) ^ 2)
                Call calculate_angle(0, 0, currentsloped, currentzp, angle)
                currentvangle = CStr(angle)
                Button1.Enabled = True
                Timer1.Enabled = False

            Case Else
                MsgBox("The total station must be set to either Topcon, Wild, Sokkia, Leica, LeicaGeo or Simulation in order to record a point. Use menu Station-Settings to modify.", MsgBoxStyle.Information And MsgBoxStyle.OkOnly, "EDM-Mobile")
        End Select

    End Sub

    Private Sub newprismname_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles newprismname.TextChanged

    End Sub

    Private Sub newprismconstant_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles newprismconstant.TextChanged

    End Sub

    Private Sub newprismheight_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles newprismheight.TextChanged

    End Sub

    Private Sub Label2_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label2.ParentChanged

    End Sub

    Private Sub Label4_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label4.ParentChanged

    End Sub

    Private Sub prismconstantlabel_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles prismconstantlabel.ParentChanged

    End Sub

    Private Sub dontwait_CheckStateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dontwait.CheckStateChanged

        If dontwait.Checked = False Then
            donotwaitforprism = False
        Else
            donotwaitforprism = True
        End If

    End Sub
End Class
