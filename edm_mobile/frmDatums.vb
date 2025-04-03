Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlServerCe
Imports System.io
Imports System
Public Class frmDatums
    Inherits System.Windows.Forms.Form
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents datumlist As System.Windows.Forms.ListBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label

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
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents datumname As System.Windows.Forms.TextBox
    Friend WithEvents datumx As System.Windows.Forms.TextBox
    Friend WithEvents datumy As System.Windows.Forms.TextBox
    Friend WithEvents datumz As System.Windows.Forms.TextBox
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.datumlist = New System.Windows.Forms.ListBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.datumname = New System.Windows.Forms.TextBox
        Me.datumx = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.datumy = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.datumz = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 113)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(123, 142)
        Me.Label1.Text = "Use this screen to edit or add datums.  You can hand enter coordinates or record " & _
            "them with the total station.  You can overwrite existing datums and you can reco" & _
            "rd new datums."
        '
        'datumlist
        '
        Me.datumlist.Location = New System.Drawing.Point(137, 8)
        Me.datumlist.Name = "datumlist"
        Me.datumlist.Size = New System.Drawing.Size(95, 142)
        Me.datumlist.TabIndex = 6
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(48, 16)
        Me.Label2.Text = "Name :"
        '
        'datumname
        '
        Me.datumname.Location = New System.Drawing.Point(56, 8)
        Me.datumname.MaxLength = 10
        Me.datumname.Name = "datumname"
        Me.datumname.Size = New System.Drawing.Size(75, 21)
        Me.datumname.TabIndex = 0
        Me.datumname.WordWrap = False
        '
        'datumx
        '
        Me.datumx.Location = New System.Drawing.Point(32, 32)
        Me.datumx.MaxLength = 15
        Me.datumx.Name = "datumx"
        Me.datumx.Size = New System.Drawing.Size(99, 21)
        Me.datumx.TabIndex = 1
        Me.datumx.WordWrap = False
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(8, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(24, 16)
        Me.Label3.Text = "X :"
        '
        'datumy
        '
        Me.datumy.Location = New System.Drawing.Point(32, 56)
        Me.datumy.MaxLength = 15
        Me.datumy.Name = "datumy"
        Me.datumy.Size = New System.Drawing.Size(99, 21)
        Me.datumy.TabIndex = 2
        Me.datumy.WordWrap = False
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 56)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(24, 16)
        Me.Label4.Text = "Y :"
        '
        'datumz
        '
        Me.datumz.Location = New System.Drawing.Point(32, 80)
        Me.datumz.MaxLength = 15
        Me.datumz.Name = "datumz"
        Me.datumz.Size = New System.Drawing.Size(99, 21)
        Me.datumz.TabIndex = 3
        Me.datumz.WordWrap = False
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 80)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(24, 16)
        Me.Label5.Text = "Z :"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(160, 167)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(72, 32)
        Me.Button1.TabIndex = 4
        Me.Button1.Text = "Record"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(160, 207)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(72, 32)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "Clear"
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.Add(Me.MenuItem1)
        Me.MainMenu1.MenuItems.Add(Me.MenuItem2)
        Me.MainMenu1.MenuItems.Add(Me.MenuItem3)
        '
        'MenuItem1
        '
        Me.MenuItem1.Text = "Save       "
        '
        'MenuItem2
        '
        Me.MenuItem2.Text = "Delete      "
        '
        'MenuItem3
        '
        Me.MenuItem3.Text = "Close"
        '
        'frmDatums
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(240, 268)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
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
        Me.Name = "frmDatums"
        Me.Text = " Datums"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        datumx.Text = ""
        datumy.Text = ""
        datumz.Text = ""
        datumname.Text = ""
    End Sub

    Private Sub frmDatums_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        populate_list()
        
    End Sub

    Private Sub datumlist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles datumlist.SelectedIndexChanged

        If datumlist.SelectedIndex = -1 Then Exit Sub
        Cursor.Current = Cursors.WaitCursor
        Dim d As String = CStr(datumlist.SelectedItem)
        Dim datums As New SqlCeConnection("Datasource=" + options(1))
        Dim sql As SqlCeCommand = datums.CreateCommand
        Try
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If (datumx.Text <> "" Or datumy.Text <> "" Or datumz.Text <> "") And datumname.Text <> "" Then
            If MsgBox("Ok to overwrite exisiting values?", MsgBoxStyle.OkCancel, "EDM-Mobile") = MsgBoxResult.Cancel Then Exit Sub
        End If

        If MsgBox("Aim at the datum and press Ok to record it.", MsgBoxStyle.OkCancel, "EDM-Mobile") = MsgBoxResult.Cancel Then Exit Sub

        'Set to a kind of X-shot and take it - this will send us back here into form.activate
        shottype = -6

        Dim whichprism As New frmWhichPrism
        If whichprism.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
            datumx.Text = CStr(currentxp)
            datumy.Text = CStr(currentyp)
            datumz.Text = CStr(currentzp)
        End If

    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click

        If Not IsNumeric(datumx.Text) Or Not IsNumeric(datumy.Text) Or Not IsNumeric(datumz.Text) Or datumname.Text = "" Then
            MessageBox.Show("Enter three coordinates and a name.", "EDM-Mobile", MessageBoxButtons.OK, MessageBoxIcon.None, MessageBoxDefaultButton.Button1)
            Exit Sub
        End If

        Dim d As String = datumname.Text
        Dim flag As Boolean = False
        Dim xp As String = datumx.Text
        Dim yp As String = datumy.Text
        Dim zp As String = datumz.Text

        Dim datums As New SqlCeConnection("Datasource='" + options(1) + "'")
        Dim sql As SqlCeCommand = datums.CreateCommand
        sql.CommandText = "SELECT name FROM edm_datums WHERE name='" + d + "'"
        Dim r As Object
        datums.Open()
        r = sql.ExecuteScalar
        If LCase(CStr(r)) = LCase(d) Then
            If MsgBox("Overwrite existing datum?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "EDM-Mobile") = MsgBoxResult.No Then Exit Sub
            Try
                sql.CommandText = "DELETE FROM edm_datums WHERE name='" + d + "'"
                sql.ExecuteNonQuery()
            Catch
                MsgBox("Error: Could not delete datum before replacing.")
            End Try
        End If

        Try
            sql.CommandText = "INSERT INTO edm_datums (name,x,y,z,day,time) VALUES ('" + d + "'," + xp + "," + yp + "," + zp + "," + Microsoft.VisualBasic.DateAndTime.DateString + ",'" + Microsoft.VisualBasic.DateAndTime.TimeString + "')"
            sql.ExecuteNonQuery()
        Catch
            MsgBox("Error: Could not add new datum.")
        End Try
        datums.Close()

        populate_list()

    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        If datumlist.SelectedIndex = -1 Then Exit Sub
        Dim d As String = datumlist.SelectedItem.ToString
        If MsgBox("Delete " + d + "?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "EDM-Mobile") <> MsgBoxResult.Yes Then Exit Sub
        Cursor.Current = Cursors.WaitCursor
        Dim datums As New SqlCeConnection("Datasource=" + options(1))
        Dim sql As SqlCeCommand = datums.CreateCommand
        Try
            datums.Open()
            sql.CommandText = "DELETE FROM edm_datums WHERE name='" + d + "'"
            sql.ExecuteNonQuery()
        Catch sqlex As System.Data.SqlServerCe.SqlCeException
            MsgBox("Error deleting datum from edm_datums in " & options(1) & vbNewLine & sqlex.Message, MsgBoxStyle.Exclamation, "Issue")
            Exit Try
        Finally
            datums.Close()
        End Try
        populate_list()
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click

        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub
    Private Sub populate_list()

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
            Exit Try

        Finally
            datums.Close()

        End Try
        datumx.Text = ""
        datumy.Text = ""
        datumz.Text = ""
        datumname.Text = ""
        datumlist.SelectedIndex = -1

    End Sub
End Class
