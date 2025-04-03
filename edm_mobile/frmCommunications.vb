Public Class frmCommunications
    Inherits System.Windows.Forms.Form
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents baudrate As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents parity As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents databits As System.Windows.Forms.ComboBox
    Friend WithEvents label40 As System.Windows.Forms.Label
    Friend WithEvents stopbits As System.Windows.Forms.ComboBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents SerialPort1 As System.IO.Ports.SerialPort
    Private components As System.ComponentModel.IContainer
    Friend WithEvents totalstation As System.Windows.Forms.ComboBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents stationhelp As System.Windows.Forms.Button
    Friend WithEvents commhelp As System.Windows.Forms.Button
    Friend WithEvents baudhelp As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label

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
    Friend WithEvents commport As System.Windows.Forms.ComboBox
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.Button1 = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.commport = New System.Windows.Forms.ComboBox
        Me.baudrate = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.parity = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.databits = New System.Windows.Forms.ComboBox
        Me.label40 = New System.Windows.Forms.Label
        Me.stopbits = New System.Windows.Forms.ComboBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.SerialPort1 = New System.IO.Ports.SerialPort(Me.components)
        Me.totalstation = New System.Windows.Forms.ComboBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.stationhelp = New System.Windows.Forms.Button
        Me.commhelp = New System.Windows.Forms.Button
        Me.baudhelp = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(8, 213)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(224, 32)
        Me.Button1.TabIndex = 11
        Me.Button1.Text = "Set"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 44)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 16)
        Me.Label1.Text = "COM Port :"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'commport
        '
        Me.commport.Location = New System.Drawing.Point(96, 44)
        Me.commport.Name = "commport"
        Me.commport.Size = New System.Drawing.Size(112, 22)
        Me.commport.TabIndex = 9
        '
        'baudrate
        '
        Me.baudrate.Items.Add("19200")
        Me.baudrate.Items.Add("9600")
        Me.baudrate.Items.Add("4800")
        Me.baudrate.Items.Add("2400")
        Me.baudrate.Items.Add("1200")
        Me.baudrate.Items.Add("600")
        Me.baudrate.Items.Add("300")
        Me.baudrate.Location = New System.Drawing.Point(96, 68)
        Me.baudrate.Name = "baudrate"
        Me.baudrate.Size = New System.Drawing.Size(112, 22)
        Me.baudrate.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 68)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(88, 16)
        Me.Label2.Text = "Baud rate :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'parity
        '
        Me.parity.Items.Add("Even")
        Me.parity.Items.Add("Odd")
        Me.parity.Items.Add("None")
        Me.parity.Location = New System.Drawing.Point(96, 92)
        Me.parity.Name = "parity"
        Me.parity.Size = New System.Drawing.Size(112, 22)
        Me.parity.TabIndex = 5
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 92)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 16)
        Me.Label3.Text = "Parity :"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'databits
        '
        Me.databits.Items.Add("8")
        Me.databits.Items.Add("7")
        Me.databits.Location = New System.Drawing.Point(96, 116)
        Me.databits.Name = "databits"
        Me.databits.Size = New System.Drawing.Size(112, 22)
        Me.databits.TabIndex = 3
        '
        'label40
        '
        Me.label40.Location = New System.Drawing.Point(8, 116)
        Me.label40.Name = "label40"
        Me.label40.Size = New System.Drawing.Size(88, 16)
        Me.label40.Text = "Data bits :"
        Me.label40.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'stopbits
        '
        Me.stopbits.Items.Add("0")
        Me.stopbits.Items.Add("1")
        Me.stopbits.Location = New System.Drawing.Point(96, 140)
        Me.stopbits.Name = "stopbits"
        Me.stopbits.Size = New System.Drawing.Size(112, 22)
        Me.stopbits.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 140)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 16)
        Me.Label4.Text = "Stop bits :"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 173)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(224, 37)
        Me.Label5.Text = "Warning: If settings are incorrect, this screen will lock for several seconds."
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.Add(Me.MenuItem1)
        '
        'MenuItem1
        '
        Me.MenuItem1.Text = "Close"
        '
        'totalstation
        '
        Me.totalstation.Items.Add("Leica")
        Me.totalstation.Items.Add("Leica-Geo")
        Me.totalstation.Items.Add("Sokkia")
        Me.totalstation.Items.Add("Topcon")
        Me.totalstation.Items.Add("Wild")
        Me.totalstation.Items.Add("Simulation")
        Me.totalstation.Items.Add("None")
        Me.totalstation.Location = New System.Drawing.Point(96, 20)
        Me.totalstation.Name = "totalstation"
        Me.totalstation.Size = New System.Drawing.Size(112, 22)
        Me.totalstation.TabIndex = 12
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(8, 20)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 16)
        Me.Label6.Text = "Total Station :"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'stationhelp
        '
        Me.stationhelp.Location = New System.Drawing.Point(213, 20)
        Me.stationhelp.Name = "stationhelp"
        Me.stationhelp.Size = New System.Drawing.Size(18, 22)
        Me.stationhelp.TabIndex = 15
        Me.stationhelp.Text = " ?"
        '
        'commhelp
        '
        Me.commhelp.Location = New System.Drawing.Point(213, 44)
        Me.commhelp.Name = "commhelp"
        Me.commhelp.Size = New System.Drawing.Size(18, 22)
        Me.commhelp.TabIndex = 16
        Me.commhelp.Text = " ?"
        '
        'baudhelp
        '
        Me.baudhelp.Location = New System.Drawing.Point(213, 68)
        Me.baudhelp.Name = "baudhelp"
        Me.baudhelp.Size = New System.Drawing.Size(18, 22)
        Me.baudhelp.TabIndex = 17
        Me.baudhelp.Text = " ?"
        '
        'frmCommunications
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(240, 268)
        Me.ControlBox = False
        Me.Controls.Add(Me.baudhelp)
        Me.Controls.Add(Me.commhelp)
        Me.Controls.Add(Me.stationhelp)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.totalstation)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.stopbits)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.databits)
        Me.Controls.Add(Me.label40)
        Me.Controls.Add(Me.parity)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.baudrate)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.commport)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Button1)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmCommunications"
        Me.Text = "Station type and COM Settings"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim p As String
        Dim iniclass As String
        Dim inidata(3, 2) As String
        Dim status As Integer

        Cursor.Current = Cursors.WaitCursor

        If commcontrol.IsOpen Then commcontrol.Close()

        options(4) = UCase(totalstation.Text)

        Select Case parity.Text
            Case "Odd"
                p = "O"
            Case "Even"
                p = "E"
            Case "None"
                p = "N"
            Case Else
                p = "N"
        End Select

        If UCase(options(4)) <> "SIMULATION" Then

            Select Case baudrate.Text
                Case "300", "600", "1200", "2400", "4800", "9600", "19200"
                Case Else
                    Cursor.Current = Cursors.Default
                    MsgBox("Invalid Baudrate setting.", MsgBoxStyle.Information, "EDM-Mobile")
                    Exit Sub
            End Select

            Select Case databits.Text
                Case "7", "8"
                Case Else
                    Cursor.Current = Cursors.Default
                    MsgBox("Invalid data bits setting.", MsgBoxStyle.Information, "EDM-Mobile")
                    Exit Sub
            End Select

            Select Case stopbits.Text
                Case "0", "1", "2"
                Case Else
                    Cursor.Current = Cursors.Default
                    MsgBox("Invalid stops bits settings.", MsgBoxStyle.Information, "EDM-Mobile")
                    Exit Sub
            End Select
        End If

        options(2) = commport.Text
        options(3) = baudrate.Text + "," + p + "," + databits.Text + "," + stopbits.Text

        open_com()

        iniclass = "[EDM-MOBILE]"
        inidata(1, 1) = "COM"
        inidata(1, 2) = options(2)
        inidata(2, 1) = "SETTINGS"
        inidata(2, 2) = options(3)
        inidata(3, 1) = "Total Station"
        inidata(3, 2) = options(4)
        Call WriteIni(edmini, iniclass, inidata, status)

        Me.DialogResult = Windows.Forms.DialogResult.OK
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub frmCommunications_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If options(2) <> "" Then
            Call set_comm_menus()
        End If

        Select Case options(4)
            Case "TOPCON"
                totalstation.Text = "Topcon"
            Case "WILD"
                totalstation.Text = "Wild"
            Case "LEICA"
                totalstation.Text = "Leica"
            Case "LEICA-GEO"
                totalstation.Text = "Leica-Geo"
            Case "SOKKIA"
                totalstation.Text = "Sokkia"
            Case "SIMULATION"
                totalstation.Text = "Simulation"
            Case "NONE"
                totalstation.Text = "None"
            Case Else
                totalstation.Text = ""
        End Select

    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click

        Me.DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub

    Private Sub set_comm_menus()

        Dim t1 As String
        Dim l(10) As String
        Dim n As Integer
        Dim s As String

        commport.Items.Clear()
        For Each s In IO.Ports.SerialPort.GetPortNames()
            commport.Items.Add(s)
        Next

        Select Case UCase(options(2))
            Case "COM1"
                commport.Text = "COM1"
            Case "COM2"
                commport.Text = "COM2"
            Case "COM3"
                commport.Text = "COM3"
            Case "COM4"
                commport.Text = "COM4"
            Case "COM5"
                commport.Text = "COM5"
            Case "COM6"
                commport.Text = "COM6"
            Case "COM7"
                commport.Text = "COM7"
            Case "COM8"
                commport.Text = "COM8"
            Case Else
                commport.Text = UCase(options(2))
        End Select

        t1 = options(3)
        If t1 <> "" Then
            Call parse_list(t1, ",", n, l)
            If l(1) <> "" Then baudrate.Text = l(1)
            If l(2) <> "" Then
                Select Case LCase(l(2))
                    Case "n"
                        parity.Text = "None"
                    Case "e"
                        parity.Text = "Even"
                    Case "o"
                        parity.Text = "Odd"
                    Case Else
                End Select
            End If
            If l(3) <> "" Then databits.Text = l(3)
            If l(4) <> "" Then stopbits.Text = l(4)
        End If

    End Sub

    Private Sub baudhelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles baudhelp.Click

        MsgBox("The baud rate, data bits, stop bits, and parity must match the settings in the total station. Refer to your total station manual to access these settings.", MsgBoxStyle.Information And MsgBoxStyle.OkOnly)

    End Sub

    Private Sub commhelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles commhelp.Click

        MsgBox("Determining the correct COM port can be difficult but normally only the available ports on this computer are included in the list.", MsgBoxStyle.Information And MsgBoxStyle.OkOnly)

    End Sub

    Private Sub stationhelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles stationhelp.Click

        MsgBox("This program has been tested with Topcon and Leica stations. It now works with newer Leica stations that use GeoCOM to communicate.  Previous versions of this code have been tested with Sokkia and Wild. Nikons are untested. Leica and Topcon - set the station to return a CR/LF. Topcon - set the station to use ACK/NACK.", MsgBoxStyle.Information And MsgBoxStyle.OkOnly)

    End Sub

End Class
