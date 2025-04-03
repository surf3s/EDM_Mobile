Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlServerCe
Imports System.io
Imports System
Public Class frmonepoint
    Inherits System.Windows.Forms.Form
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents datumlist As System.Windows.Forms.ListBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents xvalue As System.Windows.Forms.Label
    Friend WithEvents yvalue As System.Windows.Forms.Label
    Friend WithEvents zvalue As System.Windows.Forms.Label

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
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.datumlist = New System.Windows.Forms.ListBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.xvalue = New System.Windows.Forms.Label
        Me.yvalue = New System.Windows.Forms.Label
        Me.zvalue = New System.Windows.Forms.Label
        Me.Button1 = New System.Windows.Forms.Button
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(224, 46)
        Me.Label1.Text = "If horizontal angle is already set, measure a known datum to calculate current po" & _
            "sition. "
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 60)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(224, 16)
        Me.Label2.Text = "Reference datum (to be measured) :"
        '
        'datumlist
        '
        Me.datumlist.Location = New System.Drawing.Point(8, 76)
        Me.datumlist.Name = "datumlist"
        Me.datumlist.Size = New System.Drawing.Size(100, 114)
        Me.datumlist.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(120, 80)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(24, 16)
        Me.Label3.Text = "X :"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(120, 96)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(24, 16)
        Me.Label4.Text = "Y :"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(120, 114)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(24, 16)
        Me.Label5.Text = "Z :"
        '
        'xvalue
        '
        Me.xvalue.Location = New System.Drawing.Point(144, 80)
        Me.xvalue.Name = "xvalue"
        Me.xvalue.Size = New System.Drawing.Size(88, 16)
        Me.xvalue.Text = "999999.999"
        '
        'yvalue
        '
        Me.yvalue.Location = New System.Drawing.Point(144, 96)
        Me.yvalue.Name = "yvalue"
        Me.yvalue.Size = New System.Drawing.Size(88, 16)
        Me.yvalue.Text = "999999.999"
        '
        'zvalue
        '
        Me.zvalue.Location = New System.Drawing.Point(144, 114)
        Me.zvalue.Name = "zvalue"
        Me.zvalue.Size = New System.Drawing.Size(88, 16)
        Me.zvalue.Text = "999999.999"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(8, 210)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(224, 40)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Measure"
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.Add(Me.MenuItem1)
        '
        'MenuItem1
        '
        Me.MenuItem1.Text = "Close"
        '
        'frmonepoint
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(240, 270)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.zvalue)
        Me.Controls.Add(Me.yvalue)
        Me.Controls.Add(Me.xvalue)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.datumlist)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmonepoint"
        Me.Text = "EDM-Mobile"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If datumlist.SelectedIndex = -1 Then
            MsgBox("Select a datum first.  If you don't have a datum added yet, use Initialize Direct or Edit Datums to add new ones.", MsgBoxStyle.Information, "EDM-Mobile")
            Exit Sub
        End If

        Button1.Enabled = False
        MenuItem1.Enabled = False
        shottype = -3
        Dim whichprism As New frmWhichPrism
        If whichprism.ShowDialog = Windows.Forms.DialogResult.OK Then
            Dim x1 As Double = CDbl(xvalue.Text)
            Dim y1 As Double = CDbl(yvalue.Text)
            Dim z1 As Double = CDbl(zvalue.Text)
            currentstationx = CDbl(x1) - CDbl(currentxp)
            currentstationy = CDbl(y1) - CDbl(currentyp)
            currentstationz = CDbl(z1) - CDbl(currentzp)
            Call MsgBox("Current station is now " + vbCr + " X=" + CStr(currentstationx) + "," + vbCr + " Y=" + CStr(currentstationy) + "," + vbCr + " Z=" + CStr(currentstationz) + ".", MsgBoxStyle.OkOnly, "EDM-Mobile")
            currentstationname = ""
            referencedatum = datumlist.SelectedItem.ToString
            Call write_ini_header(cfgfile)

            If WriteLogFile And LogFile <> "" Then
                Dim file1 As New StreamWriter(LogFile, True)
                file1.WriteLine("Initialize one point at " & DateAndTime.Now)
                file1.WriteLine("Reference Datum : " & referencedatum)
                file1.WriteLine("Prism name : " & currentprismname)
                file1.WriteLine("Prism height : " & currentprismheight)
                file1.WriteLine("Prism offset : " & currentprismoffset)
                file1.WriteLine("Station X : " & Str(currentstationx))
                file1.WriteLine("Station Y : " & Str(currentstationy))
                file1.WriteLine("Station Z : " & Str(currentstationz))
                file1.WriteLine("")
                file1.Flush()
                file1.Close()
                file1.Dispose()
            End If

            Me.DialogResult = Windows.Forms.DialogResult.OK
        End If
        Button1.Enabled = True
        MenuItem1.Enabled = True

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

    Private Sub frmonepoint_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated

    End Sub

    Private Sub frmonepoint_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim a As Integer

        xvalue.Text = ""
        yvalue.Text = ""
        zvalue.Text = ""
        datumlist.Items.Clear()
        datumlist.ValueMember = ""

        Dim datums As New SqlCeConnection("Datasource=" + options(1))
        Dim sql As SqlCeCommand = datums.CreateCommand
        sql.CommandText = "SELECT * FROM edm_datums"

        Try
            datums.Open()
            Dim rdr As SqlCeDataReader = sql.ExecuteReader
            Do While (rdr.Read)
                If Not IsDBNull(rdr.Item("name")) Then
                    datumlist.Items.Add(rdr.Item("name"))

                    If LCase(CStr(rdr.Item("name"))) = LCase(referencedatum) Then
                        xvalue.Text = FormatNumber(rdr.Item("x"), 3)
                        yvalue.Text = FormatNumber(rdr.Item("y"), 3)
                        zvalue.Text = FormatNumber(rdr.Item("z"), 3)
                    End If

                End If
            Loop

        Catch sqlex As System.Data.SqlServerCe.SqlCeException
            MsgBox("Error reading datums table in " & options(1) & vbNewLine & sqlex.Message, MsgBoxStyle.Exclamation, "Issue")
            Exit Try

        Finally
            datums.Close()

        End Try

        For a = 0 To datumlist.Items.Count - 1
            If LCase(CStr(datumlist.Items(a))) = LCase(referencedatum) Then
                datumlist.SelectedIndex = a
                Exit For
            End If
        Next a

    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click

        Me.DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub
End Class
