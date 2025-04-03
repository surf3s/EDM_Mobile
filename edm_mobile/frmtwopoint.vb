Imports Microsoft.VisualBasic
Imports System.Math
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlServerCe
Imports System.io
Imports System
Public Class frmtwopoint
    Inherits System.Windows.Forms.Form
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents zvalue As System.Windows.Forms.Label
    Friend WithEvents yvalue As System.Windows.Forms.Label
    Friend WithEvents xvalue As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents datumlist1 As System.Windows.Forms.ListBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents datumlist2 As System.Windows.Forms.ListBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
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
    Friend WithEvents z2value As System.Windows.Forms.Label
    Friend WithEvents y2value As System.Windows.Forms.Label
    Friend WithEvents x2value As System.Windows.Forms.Label
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.zvalue = New System.Windows.Forms.Label
        Me.yvalue = New System.Windows.Forms.Label
        Me.xvalue = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.datumlist1 = New System.Windows.Forms.ListBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.z2value = New System.Windows.Forms.Label
        Me.y2value = New System.Windows.Forms.Label
        Me.x2value = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.datumlist2 = New System.Windows.Forms.ListBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 4)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(224, 45)
        Me.Label1.Text = "Place station over datum and measure a reference datum.  Horizontal angle will be" & _
            " automatically calculated and set."
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(8, 216)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(224, 40)
        Me.Button1.TabIndex = 8
        Me.Button1.Text = "Measure"
        '
        'zvalue
        '
        Me.zvalue.Location = New System.Drawing.Point(144, 100)
        Me.zvalue.Name = "zvalue"
        Me.zvalue.Size = New System.Drawing.Size(88, 16)
        Me.zvalue.Text = "999999.999"
        '
        'yvalue
        '
        Me.yvalue.Location = New System.Drawing.Point(144, 84)
        Me.yvalue.Name = "yvalue"
        Me.yvalue.Size = New System.Drawing.Size(88, 16)
        Me.yvalue.Text = "999999.999"
        '
        'xvalue
        '
        Me.xvalue.Location = New System.Drawing.Point(144, 68)
        Me.xvalue.Name = "xvalue"
        Me.xvalue.Size = New System.Drawing.Size(88, 16)
        Me.xvalue.Text = "999999.999"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(120, 100)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(24, 16)
        Me.Label5.Text = "Z :"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(120, 84)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(24, 16)
        Me.Label4.Text = "Y :"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(120, 68)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(24, 16)
        Me.Label3.Text = "X :"
        '
        'datumlist1
        '
        Me.datumlist1.Location = New System.Drawing.Point(8, 68)
        Me.datumlist1.Name = "datumlist1"
        Me.datumlist1.Size = New System.Drawing.Size(100, 58)
        Me.datumlist1.TabIndex = 15
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 52)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(224, 16)
        Me.Label2.Text = "Reference datum (to be measured) :"
        '
        'z2value
        '
        Me.z2value.Location = New System.Drawing.Point(144, 179)
        Me.z2value.Name = "z2value"
        Me.z2value.Size = New System.Drawing.Size(88, 16)
        Me.z2value.Text = "999999.999"
        '
        'y2value
        '
        Me.y2value.Location = New System.Drawing.Point(144, 163)
        Me.y2value.Name = "y2value"
        Me.y2value.Size = New System.Drawing.Size(88, 16)
        Me.y2value.Text = "999999.999"
        '
        'x2value
        '
        Me.x2value.Location = New System.Drawing.Point(144, 147)
        Me.x2value.Name = "x2value"
        Me.x2value.Size = New System.Drawing.Size(88, 16)
        Me.x2value.Text = "999999.999"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(120, 179)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(24, 16)
        Me.Label9.Text = "Z :"
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(120, 163)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(24, 16)
        Me.Label10.Text = "Y :"
        '
        'Label11
        '
        Me.Label11.Location = New System.Drawing.Point(120, 147)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(24, 16)
        Me.Label11.Text = "X :"
        '
        'datumlist2
        '
        Me.datumlist2.Location = New System.Drawing.Point(8, 147)
        Me.datumlist2.Name = "datumlist2"
        Me.datumlist2.Size = New System.Drawing.Size(100, 58)
        Me.datumlist2.TabIndex = 6
        '
        'Label12
        '
        Me.Label12.Location = New System.Drawing.Point(8, 131)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(224, 16)
        Me.Label12.Text = "Datum under station :"
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.Add(Me.MenuItem3)
        '
        'MenuItem3
        '
        Me.MenuItem3.Text = "Close"
        '
        'frmtwopoint
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
        Me.Name = "frmtwopoint"
        Me.Text = "EDM-Mobile"
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
        Catch
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
        Catch
        Finally
            datums.Close()
        End Try
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If datumlist1.SelectedIndex = -1 Or datumlist2.SelectedIndex = -1 Then
            MsgBox("Select a datum from each list.  If you need to add datums, use Initialize Direct or Edit Datums first.", MsgBoxStyle.Information, "EDMCE")
            Exit Sub
        ElseIf datumlist1.SelectedItem.ToString = datumlist2.SelectedItem.ToString Then
            MsgBox("The reference datum and the current station cannot be the same point.", MsgBoxStyle.Information, "EDMCE")
            Exit Sub
        End If

        Dim angle As String
        Dim anglenum As Double
        Dim degminsec As String
        Dim degminsectext As String
        Dim degrees As Integer
        Dim minutes As Integer
        Dim seconds As Integer
        Dim deg As Integer
        Dim min As Integer
        Dim sec As Single
        Dim x1 As Double = CDbl(x2value.Text)
        Dim y1 As Double = CDbl(y2value.Text)
        Dim z1 As Double = CDbl(z2value.Text)
        Dim x2 As Double = CDbl(xvalue.Text)
        Dim y2 As Double = CDbl(yvalue.Text)
        Dim z2 As Double = CDbl(zvalue.Text)
        Dim datum1name As String = datumlist2.SelectedItem.ToString
        Dim datum2name As String = datumlist1.SelectedItem.ToString

        Call calculate_angle(x1, y1, x2, y2, anglenum)

        Call conv_angle_to_degminsec(anglenum, degrees, minutes, seconds)

        degminsectext = CStr(degrees) + " " + Microsoft.VisualBasic.Right("00" + CStr(minutes), 2) + "' " + Microsoft.VisualBasic.Right("00" + CStr(seconds), 2) + Chr(34)
        degminsec = CStr(degrees) + "." + Microsoft.VisualBasic.Right("00" + CStr(minutes), 2) + Microsoft.VisualBasic.Right("00" + CStr(seconds), 2)

        Button1.Enabled = False
        MenuItem3.Enabled = False
        If MsgBox("Aim total station at " + datum2name + " and press OK to set angle to " + degminsectext + ".", MsgBoxStyle.OkCancel, "EDMCE") <> MsgBoxResult.Cancel Then

            angle = degminsec
            deg = 0
            min = 0
            sec = 0
            Call sethortangle(angle, deg, min, sec)

            If MsgBox("Now press Ok to record " + datum2name + ".", MsgBoxStyle.OkCancel, "EDMCE") <> MsgBoxResult.Cancel Then

                shottype = -4
                Dim whichprism As New frmWhichPrism
                If whichprism.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then

                    Dim xerror As Double = CDbl(FormatNumber(Abs(x2 - (x1 + currentxp)), 3))
                    Dim yerror As Double = CDbl(FormatNumber(Abs(y2 - (y1 + currentyp)), 3))
                    Dim zerror As Double = CDbl(FormatNumber(Abs(z2 - (z1 + currentzp)), 3))

                    If MsgBox("Error of " + Chr(13) + " X=" + CStr(xerror) + "," + Chr(13) + " Y=" + CStr(yerror) + "," + Chr(13) + " Station height of=" + CStr(zerror) + "." + Chr(13) + "Accept this station?", MsgBoxStyle.OkCancel, "EDMCE") = MsgBoxResult.Ok Then
                        currentstationx = x1
                        currentstationy = y1
                        currentstationz = z2 - currentzp
                        currentstationname = datum1name
                        referencedatum = datum2name

                        Call write_ini_header(cfgfile)

                        If WriteLogFile And LogFile <> "" Then
                            Dim file1 As New StreamWriter(LogFile, True)
                            file1.WriteLine("Initialize two point at " & DateAndTime.Now)
                            file1.WriteLine("Reference datum : " & datum2name)
                            file1.WriteLine("Reference X : " & xvalue.Text)
                            file1.WriteLine("Reference Y : " & yvalue.Text)
                            file1.WriteLine("Reference Z : " & zvalue.Text)
                            file1.WriteLine("Setup over : " & datum1name)
                            file1.WriteLine("Setup X : " & x2value.Text)
                            file1.WriteLine("Setup Y : " & y2value.Text)
                            file1.WriteLine("Refence angle : " & degminsectext)
                            file1.WriteLine("Prism name : " & currentprismname)
                            file1.WriteLine("Prism height : " & currentprismheight)
                            file1.WriteLine("Prism offset : " & currentprismoffset)
                            file1.WriteLine("Station X : " & Str(currentstationx))
                            file1.WriteLine("Station Y : " & Str(currentstationy))
                            file1.WriteLine("Station Z : " & Str(currentstationz))
                            file1.WriteLine("X error : " & Str(xerror))
                            file1.WriteLine("Y error : " & Str(yerror))
                            file1.WriteLine("Z error : " & Str(zerror))
                            file1.WriteLine("")
                            file1.Flush()
                            file1.Close()
                            file1.Dispose()
                        End If

                        Me.DialogResult = Windows.Forms.DialogResult.OK
                        Exit Sub
                    End If
                End If
            End If
            MsgBox("Warning: The horizontal angle has changed.  The station needs to be initialized.", MsgBoxStyle.Information And MsgBoxStyle.OkOnly)
        End If
        Button1.Enabled = True
        MenuItem3.Enabled = True
        
    End Sub

    Private Sub frmtwopoint_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim a As Integer

        xvalue.Text = ""
        yvalue.Text = ""
        zvalue.Text = ""
        x2value.Text = ""
        y2value.Text = ""
        z2value.Text = ""
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

        For a = 0 To datumlist1.Items.Count - 1
            If LCase(CStr(datumlist1.Items(a))) = LCase(referencedatum) Then
                datumlist1.SelectedIndex = a
                Exit For
            End If
        Next a
        For a = 0 To datumlist2.Items.Count - 1
            If LCase(CStr(datumlist2.Items(a))) = LCase(currentstationname) Then
                datumlist2.SelectedIndex = a
                Exit For
            End If
        Next a

    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click

        Me.DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub

    Private Sub xvalue_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles xvalue.ParentChanged

    End Sub
End Class
