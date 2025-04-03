Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlServerCe
Imports System.io
Imports System
Public Class frmThreePoint
    Inherits System.Windows.Forms.Form
    Friend WithEvents z2value As System.Windows.Forms.Label
    Friend WithEvents y2value As System.Windows.Forms.Label
    Friend WithEvents x2value As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents datumlist2 As System.Windows.Forms.ListBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents zvalue As System.Windows.Forms.Label
    Friend WithEvents yvalue As System.Windows.Forms.Label
    Friend WithEvents xvalue As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents datumlist1 As System.Windows.Forms.ListBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
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
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.z2value = New System.Windows.Forms.Label
        Me.y2value = New System.Windows.Forms.Label
        Me.x2value = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.datumlist2 = New System.Windows.Forms.ListBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.zvalue = New System.Windows.Forms.Label
        Me.yvalue = New System.Windows.Forms.Label
        Me.xvalue = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.datumlist1 = New System.Windows.Forms.ListBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.SuspendLayout()
        '
        'z2value
        '
        Me.z2value.Location = New System.Drawing.Point(144, 185)
        Me.z2value.Name = "z2value"
        Me.z2value.Size = New System.Drawing.Size(88, 16)
        Me.z2value.Text = "999999.999"
        '
        'y2value
        '
        Me.y2value.Location = New System.Drawing.Point(144, 169)
        Me.y2value.Name = "y2value"
        Me.y2value.Size = New System.Drawing.Size(88, 16)
        Me.y2value.Text = "999999.999"
        '
        'x2value
        '
        Me.x2value.Location = New System.Drawing.Point(144, 153)
        Me.x2value.Name = "x2value"
        Me.x2value.Size = New System.Drawing.Size(88, 16)
        Me.x2value.Text = "999999.999"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(120, 185)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(24, 16)
        Me.Label9.Text = "Z :"
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(120, 169)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(24, 16)
        Me.Label10.Text = "Y :"
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(120, 153)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(24, 16)
        Me.Label11.Text = "X :"
        '
        'datumlist2
        '
        Me.datumlist2.Location = New System.Drawing.Point(8, 153)
        Me.datumlist2.Name = "datumlist2"
        Me.datumlist2.Size = New System.Drawing.Size(100, 58)
        Me.datumlist2.TabIndex = 6
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(8, 137)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(224, 16)
        Me.Label12.Text = "Reference point 2 :"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(8, 219)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(224, 40)
        Me.Button1.TabIndex = 8
        Me.Button1.Text = "Measure"
        '
        'zvalue
        '
        Me.zvalue.Location = New System.Drawing.Point(144, 102)
        Me.zvalue.Name = "zvalue"
        Me.zvalue.Size = New System.Drawing.Size(88, 16)
        Me.zvalue.Text = "999999.999"
        '
        'yvalue
        '
        Me.yvalue.Location = New System.Drawing.Point(144, 86)
        Me.yvalue.Name = "yvalue"
        Me.yvalue.Size = New System.Drawing.Size(88, 16)
        Me.yvalue.Text = "999999.999"
        '
        'xvalue
        '
        Me.xvalue.Location = New System.Drawing.Point(144, 70)
        Me.xvalue.Name = "xvalue"
        Me.xvalue.Size = New System.Drawing.Size(88, 16)
        Me.xvalue.Text = "999999.999"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(120, 102)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(24, 16)
        Me.Label5.Text = "Z :"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(120, 86)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(24, 16)
        Me.Label4.Text = "Y :"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(120, 70)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(24, 16)
        Me.Label3.Text = "X :"
        '
        'datumlist1
        '
        Me.datumlist1.Location = New System.Drawing.Point(8, 70)
        Me.datumlist1.Name = "datumlist1"
        Me.datumlist1.Size = New System.Drawing.Size(100, 58)
        Me.datumlist1.TabIndex = 15
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 54)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(224, 16)
        Me.Label2.Text = "Reference point 1 :"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 3)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(224, 46)
        Me.Label1.Text = "Place station anywhere and measure two datums to calculate position. Use Station " & _
            "Verify to confirm."
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.Add(Me.MenuItem3)
        '
        'MenuItem3
        '
        Me.MenuItem3.Text = "Close"
        '
        'frmThreePoint
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(240, 268)
        Me.ControlBox = False
        Me.Controls.Add(Me.z2value)
        Me.Controls.Add(Me.y2value)
        Me.Controls.Add(Me.x2value)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.datumlist2)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.zvalue)
        Me.Controls.Add(Me.yvalue)
        Me.Controls.Add(Me.xvalue)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.datumlist1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmThreePoint"
        Me.Text = "Station Initialize - 3 Point"
        Me.ResumeLayout(False)

    End Sub

#End Region

 
    Private Sub datumlist1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles datumlist1.SelectedIndexChanged

        If datumlist1.SelectedIndex = -1 Then Exit Sub
        Cursor.Current = Cursors.WaitCursor
        Dim d As String = datumlist1.SelectedItem.ToString
        Dim datums As New SqlCeConnection("Datasource=" + options(1))
        Dim sql As SqlCeCommand = datums.CreateCommand
        Try
            datums.Open()
            sql.CommandText = "SELECT * FROM edm_datums WHERE name='" + d + "'"
            Dim rdr As SqlCeDataReader = sql.ExecuteReader
            Do While rdr.Read
                xvalue.Text = Format(rdr.Item("x"), "########0.000")
                yvalue.Text = Format(rdr.Item("y"), "########0.000")
                zvalue.Text = Format(rdr.Item("z"), "########0.000")
            Loop
        Catch sqlex As System.Data.SqlServerCe.SqlCeException
            MsgBox("Error reading from edm_datums in " & options(1) & vbNewLine & sqlex.Message, MsgBoxStyle.Exclamation, "Issue")
            Exit Try
        Finally
            datums.Close()
        End Try
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub datumlist2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles datumlist2.SelectedIndexChanged

        If datumlist2.SelectedIndex = -1 Then Exit Sub
        Cursor.Current = Cursors.WaitCursor
        Dim d As String = datumlist2.SelectedItem.ToString
        Dim datums As New SqlCeConnection("Datasource=" + options(1))
        Dim sql As SqlCeCommand = datums.CreateCommand
        Try
            datums.Open()
            sql.CommandText = "SELECT * FROM edm_datums WHERE name='" + d + "'"
            Dim rdr As SqlCeDataReader = sql.ExecuteReader
            Do While rdr.Read
                x2value.Text = Format(rdr.Item("x"), "########0.000")
                y2value.Text = Format(rdr.Item("y"), "########0.000")
                z2value.Text = Format(rdr.Item("z"), "########0.000")
            Loop
        Catch sqlex As System.Data.SqlServerCe.SqlCeException
            MsgBox("Error reading from edm_datums in " & options(1) & vbNewLine & sqlex.Message, MsgBoxStyle.Exclamation, "Issue")
            Exit Try
        Finally
            datums.Close()
        End Try
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub frmThreePoint_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim a As Integer

        xvalue.Text = ""
        yvalue.Text = ""
        zvalue.Text = ""
        x2value.Text = ""
        y2value.Text = ""
        z2value.Text = ""
        datumlist1.BeginUpdate()
        datumlist2.BeginUpdate()
        datumlist1.Items.Clear()
        datumlist2.Items.Clear()
        datumlist1.ValueMember = ""
        datumlist2.ValueMember = ""

        Dim datums As New SqlCeConnection("Datasource=" + options(1))
        Dim sql As SqlCeCommand = datums.CreateCommand
        sql.CommandText = "SELECT * FROM edm_datums"

        Try
            datums.Open()
            Dim rdr As SqlCeDataReader = sql.ExecuteReader
            Do While (rdr.Read)
                If Not IsDBNull(rdr.Item("name")) Then
                    datumlist1.Items.Add(rdr.Item("name"))

                    If LCase(CStr(rdr.Item("name"))) = LCase(referencedatum) Then
                        xvalue.Text = FormatNumber(rdr.Item("x"), 3)
                        yvalue.Text = FormatNumber(rdr.Item("y"), 3)
                        zvalue.Text = FormatNumber(rdr.Item("z"), 3)
                    End If

                    datumlist2.Items.Add(rdr.Item("name"))
                    If LCase(CStr(rdr.Item("name"))) = LCase(currentstationname) Then
                        x2value.Text = FormatNumber(rdr.Item("x"), 3)
                        y2value.Text = FormatNumber(rdr.Item("y"), 3)
                        z2value.Text = FormatNumber(rdr.Item("z"), 3)
                    End If

                End If
            Loop

        Catch sqlex As System.Data.SqlServerCe.SqlCeException
            MsgBox("Error reading datums table in " & options(1) & vbNewLine & sqlex.Message, MsgBoxStyle.Exclamation, "Issue")
            Exit Try

        Finally
            datums.Close()

        End Try

        Dim iniclass As String = "[EDM]"
        Dim inidata(2, 2) As String
        Dim status As Integer = 0
        inidata(1, 1) = "Referencedatum"
        inidata(1, 2) = ""
        inidata(2, 1) = "Referencedatum2"
        inidata(2, 2) = ""
        Call ReadIni(cfgfile, iniclass, inidata, status)
        If inidata(1, 2) <> "" Then referencedatum = inidata(1, 2)
        If inidata(2, 2) <> "" Then referencedatum2 = inidata(2, 2)

        For a = 0 To datumlist1.Items.Count - 1
            If LCase(CStr(datumlist1.Items(a))) = LCase(referencedatum) Then
                datumlist1.SelectedIndex = a
                Exit For
            End If
        Next a
        For a = 0 To datumlist2.Items.Count - 1
            If LCase(CStr(datumlist2.Items(a))) = LCase(referencedatum2) Then
                datumlist2.SelectedIndex = a
                Exit For
            End If
        Next a

        datumlist1.EndUpdate()
        datumlist2.EndUpdate()

    End Sub

    Private Sub frmThreePoint_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If datumlist1.SelectedIndex = -1 Or datumlist2.SelectedIndex = -1 Then
            MsgBox("Select a datum from each list.", MsgBoxStyle.Information, "EDM-Mobile")
            Exit Sub
        End If

        If datumlist1.SelectedItem.ToString = datumlist2.SelectedItem.ToString Then
            MsgBox("Select two different datums.", MsgBoxStyle.Information, "EDM-Mobile")
            Exit Sub
        End If

        Dim datum1name As String = datumlist1.SelectedItem.ToString
        Dim datum2name As String = datumlist2.SelectedItem.ToString

        Button1.Enabled = False
        MenuItem3.Enabled = False

        If MsgBox("Aim at " + datum1name + " and press Ok to record it.", MsgBoxStyle.OkCancel, "EDM-Mobile") <> MsgBoxResult.Cancel Then

            Cursor.Current = Cursors.WaitCursor

            'First, set the angle to zero on this shot
            Dim angle As String = ""
            Call sethortangle(angle, 0, 0, 0)

            Dim whichprism As New frmWhichPrism
            shottype = -5
            If whichprism.ShowDialog = Windows.Forms.DialogResult.OK Then

                Dim firstprismname As String = currentprismname
                Dim firstprismheight As Single = currentprismheight
                Dim firstprismoffset As Single = currentprismoffset

                firstxp = currentxp
                firstyp = currentyp
                firstzp = currentzp

                If MsgBox("Now aim at " + datum2name + " and press Ok to record it.", MsgBoxStyle.OkCancel, "EDM-Mobile") <> MsgBoxResult.Cancel Then

                    If whichprism.ShowDialog = Windows.Forms.DialogResult.OK Then

                        Cursor.Current = Cursors.WaitCursor

                        Dim degrees As Integer
                        Dim minutes As Integer
                        Dim seconds As Integer
                        Dim defineddistance As Double
                        Dim measureddistance As Double
                        Dim measuredangle As Double
                        Dim definedangle As Double
                        Dim datum1x As Double = CDbl(xvalue.Text)
                        Dim datum1y As Double = CDbl(yvalue.Text)
                        Dim datum1z As Double = CDbl(zvalue.Text)
                        Dim datum2x As Double = CDbl(x2value.Text)
                        Dim datum2y As Double = CDbl(y2value.Text)
                        Dim datum2z As Double = CDbl(z2value.Text)
                        Dim secondprismname As String = currentprismname
                        Dim secondprismheight As Single = currentprismheight
                        Dim secondprismoffset As Single = currentprismoffset
                        Dim info As String

                        'MsgBox(datum1x & "," & datum1y & "," & datum1z & "," & datum2x & "," & datum2y & "," & datum2z)
                        'MsgBox(firstxp & "," & firstyp & "," & firstzp & "," & currentxp & "," & currentyp & "," & currentzp)

                        defineddistance = System.Math.Sqrt((datum1x - datum2x) ^ 2 + (datum1y - datum2y) ^ 2 + (datum1z - datum2z) ^ 2)
                        measureddistance = System.Math.Sqrt((firstxp - currentxp) ^ 2 + (firstyp - currentyp) ^ 2 + (firstzp - currentzp) ^ 2)
                        If System.Math.Abs(defineddistance - measureddistance) > 0.01 Then
                            info = " Warning: The measured distance " & FormatNumber(measureddistance, 3) & vbCr
                            info = info & " is more than 1cm more than the defined distance " & FormatNumber(defineddistance, 3) & vbCr
                            Cursor.Current = Cursors.Default
                            MsgBox(info, MsgBoxStyle.Information And MsgBoxStyle.OkOnly, "EDM-Mobile")
                            Cursor.Current = Cursors.WaitCursor
                        End If

                        firstyp = -firstyp
                        'Call calculate_angle(currentxp, currentyp + firstyp, 0, 0, measuredangle)
                        'Call calculate_angle(datum2x, datum2y, datum1x, datum1y, definedangle)
                        Call calculate_angle(0, 0, currentxp, currentyp + firstyp, measuredangle)
                        Call calculate_angle(datum1x, datum1y, datum2x, datum2y, definedangle)

                        Dim angledifference As Double = definedangle - measuredangle
                        Dim cosangle As Double = System.Math.Cos(angledifference * 0.0174532925199433)
                        Dim sinangle As Double = System.Math.Sin(angledifference * 0.0174532925199433)

                        currentstationx = firstyp * sinangle + datum1x
                        currentstationy = firstyp * cosangle + datum1y
                        currentstationz = ((datum1z - firstzp) + (datum2z - currentzp)) / 2
                        currentstationx = CDbl(FormatNumber(currentstationx, 3))
                        currentstationy = CDbl(FormatNumber(currentstationy, 3))
                        currentstationz = CDbl(FormatNumber(currentstationz, 3))

                        Dim foresightsb As Double
                        Call calculate_angle(currentstationx, currentstationy, datum2x, datum2y, foresightsb)
                        Call conv_angle_to_degminsec(foresightsb, degrees, minutes, seconds)

                        Dim angleinfo As String = FormatNumber(degrees, 3) & "." & FormatNumber(minutes, 3) & " " & FormatNumber(seconds, 3)

                        Call sethortangle("", degrees, minutes, seconds)

                        currentstationname = ""
                        referencedatum = datumlist1.Text
                        referencedatum2 = datumlist2.Text

                        info = " Station X = " & FormatNumber(currentstationx, 3) & Chr(13)
                        info = info & " Station Y = " & FormatNumber(currentstationy, 3) & Chr(13)
                        info = info & " Station Z = " & FormatNumber(currentstationz, 3) & Chr(13)

                        If WriteLogFile And LogFile <> "" Then
                            Dim file1 As New StreamWriter(LogFile, True)
                            file1.WriteLine("Initialize three point at " & DateAndTime.Now)
                            file1.WriteLine("Datum 1 : " & datum1name)
                            file1.WriteLine("Datum 1 X : " & xvalue.Text)
                            file1.WriteLine("Datum 1 Y : " & yvalue.Text)
                            file1.WriteLine("Datum 1 Z : " & zvalue.Text)
                            file1.WriteLine("Prism name : " & firstprismname)
                            file1.WriteLine("Prism height : " & firstprismheight)
                            file1.WriteLine("Prism offset : " & firstprismoffset)
                            file1.WriteLine("Datum 2 : " & datum2name)
                            file1.WriteLine("Datum 2 X : " & x2value.Text)
                            file1.WriteLine("Datum 2 Y : " & y2value.Text)
                            file1.WriteLine("Datum 2 Z : " & z2value.Text)
                            file1.WriteLine("Prism name : " & secondprismname)
                            file1.WriteLine("Prism height : " & secondprismheight)
                            file1.WriteLine("Prism offset : " & secondprismoffset)
                            file1.WriteLine("Angle set to : " & angleinfo)
                            file1.WriteLine("Station X : " & Str(currentstationx))
                            file1.WriteLine("Station Y : " & Str(currentstationy))
                            file1.WriteLine("Station Z : " & Str(currentstationz))
                            file1.WriteLine("")
                            file1.Flush()
                            file1.Close()
                            file1.Dispose()
                        End If

                        Call write_ini_header(cfgfile)

                        Cursor.Current = Cursors.Default
                        MsgBox(info, MsgBoxStyle.Information And MsgBoxStyle.OkOnly, "EDM-Mobile")

                        Me.DialogResult = Windows.Forms.DialogResult.OK
                        Exit Sub

                    End If
                End If
            End If
            MsgBox("Warning: The horizontal angle has been changed. The station needs to be initialized.", MsgBoxStyle.Information And MsgBoxStyle.OkOnly)
        End If
        Button1.Enabled = True
        MenuItem3.Enabled = True

    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click

        Me.DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub
End Class
