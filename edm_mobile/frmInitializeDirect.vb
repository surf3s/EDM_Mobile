Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlServerCe
Imports System.io
Imports System
Public Class frmInitializeDirect
    Inherits System.Windows.Forms.Form
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents datumz As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents datumy As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents datumx As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents datumname As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents datumlist As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label

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
    Friend WithEvents stationheight As System.Windows.Forms.TextBox
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Button2 = New System.Windows.Forms.Button
        Me.datumz = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.datumy = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.datumx = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.datumname = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.datumlist = New System.Windows.Forms.ListBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.stationheight = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.SuspendLayout()
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(82, 81)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(72, 25)
        Me.Button2.TabIndex = 4
        Me.Button2.Text = "Clear"
        '
        'datumz
        '
        Me.datumz.Location = New System.Drawing.Point(42, 55)
        Me.datumz.Name = "datumz"
        Me.datumz.Size = New System.Drawing.Size(112, 21)
        Me.datumz.TabIndex = 5
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(18, 55)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(24, 16)
        Me.Label5.Text = "Z :"
        '
        'datumy
        '
        Me.datumy.Location = New System.Drawing.Point(42, 31)
        Me.datumy.Name = "datumy"
        Me.datumy.Size = New System.Drawing.Size(112, 21)
        Me.datumy.TabIndex = 7
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(18, 31)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(24, 16)
        Me.Label4.Text = "Y :"
        '
        'datumx
        '
        Me.datumx.Location = New System.Drawing.Point(42, 7)
        Me.datumx.Name = "datumx"
        Me.datumx.Size = New System.Drawing.Size(112, 21)
        Me.datumx.TabIndex = 9
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(18, 7)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(24, 16)
        Me.Label3.Text = "X :"
        '
        'datumname
        '
        Me.datumname.Location = New System.Drawing.Point(66, 134)
        Me.datumname.Name = "datumname"
        Me.datumname.Size = New System.Drawing.Size(88, 21)
        Me.datumname.TabIndex = 11
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(18, 134)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 16)
        Me.Label2.Text = "Name :"
        '
        'datumlist
        '
        Me.datumlist.Location = New System.Drawing.Point(160, 7)
        Me.datumlist.Name = "datumlist"
        Me.datumlist.Size = New System.Drawing.Size(72, 100)
        Me.datumlist.TabIndex = 13
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(10, 169)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(222, 79)
        Me.Label1.Text = "Use this method if the horizontal angle is already set, the station is over a kno" & _
            "wn point, and no other points are known.  Be sure to enter the station height."
        '
        'stationheight
        '
        Me.stationheight.Location = New System.Drawing.Point(66, 110)
        Me.stationheight.Name = "stationheight"
        Me.stationheight.Size = New System.Drawing.Size(88, 21)
        Me.stationheight.TabIndex = 2
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(10, 110)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(48, 16)
        Me.Label6.Text = "Height :"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(160, 115)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(64, 16)
        Me.Label7.Text = "(optional)"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(160, 139)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(64, 16)
        Me.Label8.Text = "(optional)"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.Add(Me.MenuItem1)
        Me.MainMenu1.MenuItems.Add(Me.MenuItem2)
        '
        'MenuItem1
        '
        Me.MenuItem1.Text = "Initialize"
        '
        'MenuItem2
        '
        Me.MenuItem2.Text = "Cancel"
        '
        'frmInitializeDirect
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(240, 268)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.stationheight)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.datumz)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.datumy)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.datumx)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.datumname)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.datumlist)
        Me.Controls.Add(Me.Label1)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmInitializeDirect"
        Me.Text = "Initialize Station - Direct"
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub frmInitializeDirect_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Cursor.Current = Cursors.WaitCursor
        datumlist.Items.Clear()
        Dim datums As New SqlCeConnection("Datasource=" + options(1))
        Dim sql As SqlCeCommand = datums.CreateCommand
        sql.CommandText = "SELECT * FROM edm_datums"

        Try
            datumlist.Items.Clear()
            datums.Open()
            Dim rdr As SqlCeDataReader = sql.ExecuteReader
            Do While (rdr.Read)
                datumlist.Items.Add(rdr.Item("name"))
            Loop

        Catch sqlex As System.Data.SqlServerCe.SqlCeException
            MsgBox("Error reading datums table in " & options(1) & vbNewLine & sqlex.Message, MsgBoxStyle.Exclamation, "Issue")
            Exit Sub

        Finally
            datums.Close()

        End Try

        datumx.Text = ""
        datumy.Text = ""
        datumz.Text = ""
        datumname.Text = ""
        datumlist.SelectedIndex = -1

        Dim a As Integer
        For a = 0 To datumlist.Items.Count - 1
            If LCase(datumlist.Items(a).ToString) = LCase(currentstationname) Then
                datumlist.SelectedItem = a
                Exit For
            End If
        Next a
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        datumx.Text = ""
        datumy.Text = ""
        datumz.Text = ""
        stationheight.Text = ""
        datumname.Text = ""

    End Sub

    Private Sub datumlist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles datumlist.SelectedIndexChanged

        If datumlist.SelectedIndex = -1 Then Exit Sub
        Cursor.Current = Cursors.WaitCursor
        Dim d As String
        Dim datums As New SqlCeConnection("Datasource=" + options(1))
        Dim sql As SqlCeCommand = datums.CreateCommand
        Try
            d = datumlist.SelectedItem.ToString
            datums.Open()
            sql.CommandText = "SELECT * FROM edm_datums WHERE name='" + d + "'"
            Dim rdr As SqlCeDataReader = sql.ExecuteReader
            Do While rdr.Read
                datumx.Text = Format(rdr.Item("x"), "########0.000")
                datumy.Text = Format(rdr.Item("y"), "########0.000")
                datumz.Text = Format(rdr.Item("z"), "########0.000")
                datumname.Text = rdr.Item("name").ToString
            Loop
        Catch sqlex As System.Data.SqlServerCe.SqlCeException
            MsgBox("Error reading from edm_datums in " & options(1) & vbNewLine & sqlex.Message, MsgBoxStyle.Exclamation, "Issue")
            Exit Try
        Finally
            datums.Close()
        End Try
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub frmInitializeDirect_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click

        If datumx.Text = "" Or datumy.Text = "" Or datumz.Text = "" Then
            MsgBox("Enter valid numeric coordinates", MsgBoxStyle.Information, "EDM-Mobile")
            Exit Sub
        End If

        If Not IsNumeric(datumx.Text) Or Not IsNumeric(datumy.Text) Or Not IsNumeric(datumz.Text) Then
            MsgBox("Enter valid numeric coordinates", MsgBoxStyle.Information, "EDM-Mobile")
            Exit Sub
        End If

        If stationheight.Text <> "" Then
            If Not IsNumeric(stationheight.Text) Then
                MsgBox("Enter a valid numeric station height (meters or feet) or none at all.", MsgBoxStyle.Information, "EDM-Mobile")
            End If
        End If

        Cursor.Current = Cursors.WaitCursor

        If datumx.Text <> "" And datumy.Text <> "" And datumz.Text <> "" Then
            currentstationx = CDbl(datumx.Text)
            currentstationy = CDbl(datumy.Text)
            currentstationz = CDbl(datumz.Text)
            If stationheight.Text <> "" Then
                currentstationz = currentstationz + CDbl(stationheight.Text)
            End If
            If datumname.Text <> "" Then
                currentstationname = datumname.Text

                'Check to see if this name is in use already, otherwise add it to the datum list
                Dim datums As New SqlCeConnection("Datasource='" + options(1) + "'")
                Dim sql As SqlCeCommand = datums.CreateCommand
                sql.CommandText = "SELECT name FROM edm_datums WHERE name='" + datumname.Text + "'"
                Dim r As Object
                datums.Open()
                r = sql.ExecuteScalar
                If LCase(CStr(r)) <> LCase(datumname.Text) Then
                    sql.CommandText = "INSERT INTO edm_datums (name,x,y,z,day,time) VALUES ('" + datumname.Text + "'," + datumx.Text + "," + datumy.Text + "," + datumz.Text + ",'" + Microsoft.VisualBasic.DateAndTime.DateString + "','" + Microsoft.VisualBasic.DateAndTime.TimeString + "')"
                    sql.ExecuteNonQuery()
                End If
                datums.Close()

            Else
                currentstationname = ""
            End If

        Else
            MsgBox("Select or enter datum information first.", vbOKOnly, "EDM-Mobile")
            Exit Sub

        End If

        referencedatum = ""
        Call write_ini_header(cfgfile)

        If WriteLogFile And LogFile <> "" Then
            Dim file1 As New StreamWriter(LogFile, True)
            file1.WriteLine("Initialize direct at " & DateAndTime.Now)
            file1.WriteLine("Station Name : " & currentstationname)
            file1.WriteLine("Station X : " & Str(currentstationx))
            file1.WriteLine("Station Y : " & Str(currentstationy))
            file1.WriteLine("Station Z : " & Str(currentstationz))
            file1.WriteLine("Station Height : " & stationheight.Text)
            file1.WriteLine("")
            file1.Flush()
            file1.Close()
            file1.Dispose()
        End If

        Cursor.Current = Cursors.Default
        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        Me.DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub
End Class
