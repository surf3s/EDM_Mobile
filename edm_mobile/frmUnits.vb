Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlServerCe
Imports System.io
Imports System
Public Class frmUnits
    Inherits System.Windows.Forms.Form
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents Page1 As System.Windows.Forms.TabPage
    Friend WithEvents unitname As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents unitsgrid As System.Windows.Forms.DataGrid
    Friend WithEvents bs1 As System.Windows.Forms.BindingSource
    Private components As System.ComponentModel.IContainer
    Friend WithEvents txtunitfields As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
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
    Friend WithEvents Page2 As System.Windows.Forms.TabPage
    Friend WithEvents Page3 As System.Windows.Forms.TabPage
    Friend WithEvents fieldlist As System.Windows.Forms.ListBox
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents unitfieldlist As System.Windows.Forms.ListBox
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents Button7 As System.Windows.Forms.Button
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents unitgrid As System.Windows.Forms.DataGrid
    Friend WithEvents unitlist As System.Windows.Forms.ListBox
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents rectangle As System.Windows.Forms.RadioButton
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents radius As System.Windows.Forms.TextBox
    Friend WithEvents miny As System.Windows.Forms.TextBox
    Friend WithEvents minyasd As System.Windows.Forms.Label
    Friend WithEvents maxx As System.Windows.Forms.TextBox
    Friend WithEvents minx As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents centery As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents centerx As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents maxy As System.Windows.Forms.TextBox
    Friend WithEvents circle As System.Windows.Forms.RadioButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.Page1 = New System.Windows.Forms.TabPage
        Me.rectangle = New System.Windows.Forms.RadioButton
        Me.Label7 = New System.Windows.Forms.Label
        Me.radius = New System.Windows.Forms.TextBox
        Me.miny = New System.Windows.Forms.TextBox
        Me.minyasd = New System.Windows.Forms.Label
        Me.maxx = New System.Windows.Forms.TextBox
        Me.minx = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.centery = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.centerx = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.maxy = New System.Windows.Forms.TextBox
        Me.circle = New System.Windows.Forms.RadioButton
        Me.unitname = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.unitlist = New System.Windows.Forms.ListBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Page2 = New System.Windows.Forms.TabPage
        Me.unitsgrid = New System.Windows.Forms.DataGrid
        Me.unitgrid = New System.Windows.Forms.DataGrid
        Me.Page3 = New System.Windows.Forms.TabPage
        Me.Label10 = New System.Windows.Forms.Label
        Me.Button7 = New System.Windows.Forms.Button
        Me.Button6 = New System.Windows.Forms.Button
        Me.unitfieldlist = New System.Windows.Forms.ListBox
        Me.fieldlist = New System.Windows.Forms.ListBox
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.bs1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.txtunitfields = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TabControl1.SuspendLayout()
        Me.Page1.SuspendLayout()
        Me.Page2.SuspendLayout()
        Me.Page3.SuspendLayout()
        CType(Me.bs1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.Page1)
        Me.TabControl1.Controls.Add(Me.Page2)
        Me.TabControl1.Controls.Add(Me.Page3)
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(240, 264)
        Me.TabControl1.TabIndex = 0
        '
        'Page1
        '
        Me.Page1.Controls.Add(Me.rectangle)
        Me.Page1.Controls.Add(Me.Label7)
        Me.Page1.Controls.Add(Me.radius)
        Me.Page1.Controls.Add(Me.miny)
        Me.Page1.Controls.Add(Me.minyasd)
        Me.Page1.Controls.Add(Me.maxx)
        Me.Page1.Controls.Add(Me.minx)
        Me.Page1.Controls.Add(Me.Label3)
        Me.Page1.Controls.Add(Me.centery)
        Me.Page1.Controls.Add(Me.Label4)
        Me.Page1.Controls.Add(Me.Label8)
        Me.Page1.Controls.Add(Me.centerx)
        Me.Page1.Controls.Add(Me.Label5)
        Me.Page1.Controls.Add(Me.Label9)
        Me.Page1.Controls.Add(Me.maxy)
        Me.Page1.Controls.Add(Me.circle)
        Me.Page1.Controls.Add(Me.unitname)
        Me.Page1.Controls.Add(Me.Label2)
        Me.Page1.Controls.Add(Me.unitlist)
        Me.Page1.Controls.Add(Me.Label1)
        Me.Page1.Location = New System.Drawing.Point(0, 0)
        Me.Page1.Name = "Page1"
        Me.Page1.Size = New System.Drawing.Size(240, 241)
        Me.Page1.Text = "Limits"
        '
        'rectangle
        '
        Me.rectangle.Checked = True
        Me.rectangle.Location = New System.Drawing.Point(145, 3)
        Me.rectangle.Name = "rectangle"
        Me.rectangle.Size = New System.Drawing.Size(80, 16)
        Me.rectangle.TabIndex = 0
        Me.rectangle.Text = "Rectangle"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(121, 99)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(48, 16)
        Me.Label7.Text = "Radius"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'radius
        '
        Me.radius.Location = New System.Drawing.Point(169, 99)
        Me.radius.Name = "radius"
        Me.radius.Size = New System.Drawing.Size(64, 21)
        Me.radius.TabIndex = 7
        '
        'miny
        '
        Me.miny.Location = New System.Drawing.Point(47, 51)
        Me.miny.Name = "miny"
        Me.miny.Size = New System.Drawing.Size(64, 21)
        Me.miny.TabIndex = 2
        '
        'minyasd
        '
        Me.minyasd.Location = New System.Drawing.Point(7, 51)
        Me.minyasd.Name = "minyasd"
        Me.minyasd.Size = New System.Drawing.Size(40, 16)
        Me.minyasd.Text = "Min Y:"
        '
        'maxx
        '
        Me.maxx.Location = New System.Drawing.Point(47, 75)
        Me.maxx.Name = "maxx"
        Me.maxx.Size = New System.Drawing.Size(64, 21)
        Me.maxx.TabIndex = 3
        '
        'minx
        '
        Me.minx.Location = New System.Drawing.Point(47, 27)
        Me.minx.Name = "minx"
        Me.minx.Size = New System.Drawing.Size(64, 21)
        Me.minx.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(7, 27)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 16)
        Me.Label3.Text = "Min X:"
        '
        'centery
        '
        Me.centery.Location = New System.Drawing.Point(169, 75)
        Me.centery.Name = "centery"
        Me.centery.Size = New System.Drawing.Size(64, 21)
        Me.centery.TabIndex = 6
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(7, 75)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(48, 16)
        Me.Label4.Text = "Max X:"
        '
        'Label8
        '
        Me.Label8.Location = New System.Drawing.Point(153, 75)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(16, 16)
        Me.Label8.Text = "Y"
        '
        'centerx
        '
        Me.centerx.Location = New System.Drawing.Point(169, 51)
        Me.centerx.Name = "centerx"
        Me.centerx.Size = New System.Drawing.Size(64, 21)
        Me.centerx.TabIndex = 5
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(7, 99)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(40, 16)
        Me.Label5.Text = "Max Y:"
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(153, 51)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(24, 16)
        Me.Label9.Text = "X"
        '
        'maxy
        '
        Me.maxy.Location = New System.Drawing.Point(47, 99)
        Me.maxy.Name = "maxy"
        Me.maxy.Size = New System.Drawing.Size(64, 21)
        Me.maxy.TabIndex = 4
        '
        'circle
        '
        Me.circle.Location = New System.Drawing.Point(145, 19)
        Me.circle.Name = "circle"
        Me.circle.Size = New System.Drawing.Size(56, 16)
        Me.circle.TabIndex = 15
        Me.circle.Text = "Cirlce"
        '
        'unitname
        '
        Me.unitname.Location = New System.Drawing.Point(47, 3)
        Me.unitname.Name = "unitname"
        Me.unitname.Size = New System.Drawing.Size(84, 21)
        Me.unitname.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(7, 3)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 16)
        Me.Label2.Text = "Name:"
        '
        'unitlist
        '
        Me.unitlist.Location = New System.Drawing.Point(137, 132)
        Me.unitlist.Name = "unitlist"
        Me.unitlist.Size = New System.Drawing.Size(96, 100)
        Me.unitlist.TabIndex = 18
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(11, 128)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(120, 104)
        Me.Label1.Text = "Units retain the last values for each unit.  After a shot, EDM-Mobile automatical" & _
            "ly finds the correct unit and loads the last values for that unit."
        '
        'Page2
        '
        Me.Page2.Controls.Add(Me.unitsgrid)
        Me.Page2.Controls.Add(Me.unitgrid)
        Me.Page2.Location = New System.Drawing.Point(0, 0)
        Me.Page2.Name = "Page2"
        Me.Page2.Size = New System.Drawing.Size(232, 238)
        Me.Page2.Text = "Grid View"
        '
        'unitsgrid
        '
        Me.unitsgrid.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.unitsgrid.Location = New System.Drawing.Point(0, 0)
        Me.unitsgrid.Name = "unitsgrid"
        Me.unitsgrid.Size = New System.Drawing.Size(232, 241)
        Me.unitsgrid.TabIndex = 1
        '
        'unitgrid
        '
        Me.unitgrid.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.unitgrid.Location = New System.Drawing.Point(0, 0)
        Me.unitgrid.Name = "unitgrid"
        Me.unitgrid.Size = New System.Drawing.Size(232, 248)
        Me.unitgrid.TabIndex = 0
        '
        'Page3
        '
        Me.Page3.Controls.Add(Me.Label6)
        Me.Page3.Controls.Add(Me.txtunitfields)
        Me.Page3.Controls.Add(Me.Label10)
        Me.Page3.Controls.Add(Me.Button7)
        Me.Page3.Controls.Add(Me.Button6)
        Me.Page3.Controls.Add(Me.unitfieldlist)
        Me.Page3.Controls.Add(Me.fieldlist)
        Me.Page3.Location = New System.Drawing.Point(0, 0)
        Me.Page3.Name = "Page3"
        Me.Page3.Size = New System.Drawing.Size(240, 241)
        Me.Page3.Text = "Associated Fields"
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(16, 8)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(208, 48)
        Me.Label10.Text = "Specify here the fields that should be retained for each excavation unit (eg. lev" & _
            "el, excavator, etc.)."
        '
        'Button7
        '
        Me.Button7.Location = New System.Drawing.Point(104, 146)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(32, 32)
        Me.Button7.TabIndex = 1
        Me.Button7.Text = "<<"
        '
        'Button6
        '
        Me.Button6.Location = New System.Drawing.Point(104, 106)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(32, 32)
        Me.Button6.TabIndex = 2
        Me.Button6.Text = ">>"
        '
        'unitfieldlist
        '
        Me.unitfieldlist.Location = New System.Drawing.Point(144, 64)
        Me.unitfieldlist.Name = "unitfieldlist"
        Me.unitfieldlist.Size = New System.Drawing.Size(72, 114)
        Me.unitfieldlist.TabIndex = 5
        '
        'fieldlist
        '
        Me.fieldlist.Location = New System.Drawing.Point(24, 64)
        Me.fieldlist.Name = "fieldlist"
        Me.fieldlist.Size = New System.Drawing.Size(72, 114)
        Me.fieldlist.TabIndex = 6
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.Add(Me.MenuItem3)
        Me.MainMenu1.MenuItems.Add(Me.MenuItem1)
        Me.MainMenu1.MenuItems.Add(Me.MenuItem2)
        '
        'MenuItem3
        '
        Me.MenuItem3.Text = "Save         "
        '
        'MenuItem1
        '
        Me.MenuItem1.Text = "Delete      "
        '
        'MenuItem2
        '
        Me.MenuItem2.Text = "Close"
        '
        'txtunitfields
        '
        Me.txtunitfields.Location = New System.Drawing.Point(24, 197)
        Me.txtunitfields.Multiline = True
        Me.txtunitfields.Name = "txtunitfields"
        Me.txtunitfields.Size = New System.Drawing.Size(192, 41)
        Me.txtunitfields.TabIndex = 7
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(24, 181)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(209, 13)
        Me.Label6.Text = "Current list of Associated Unit Fields"
        '
        'frmUnits
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(240, 268)
        Me.ControlBox = False
        Me.Controls.Add(Me.TabControl1)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmUnits"
        Me.Text = "Define Excavation Units"
        Me.TabControl1.ResumeLayout(False)
        Me.Page1.ResumeLayout(False)
        Me.Page2.ResumeLayout(False)
        Me.Page3.ResumeLayout(False)
        CType(Me.bs1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        Dim status As Integer
        Dim a As Integer
        options(17) = "UNIT,ID"
        For a = 0 To unitfieldlist.Items.Count - 1
            options(17) = options(17) + "," + unitfieldlist.Items(a).ToString
        Next a
        Dim iniclass As String
        Dim inidata(1, 2) As String
        iniclass = "[EDM]"
        inidata(1, 1) = "Unitfields"
        inidata(1, 2) = options(17)
        Call WriteIni(cfgfile, iniclass, inidata, status)

        Me.DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
    Private Sub populate_list()

        Dim units As New SqlCeConnection("datasource='" + options(1) + "'")
        Dim sql As SqlCeCommand = units.CreateCommand

        Try
            units.Open()
            sql.CommandText = "SELECT unit FROM edm_units"
            Dim rdr As SqlCeDataReader = sql.ExecuteReader
            unitlist.Items.Clear()
            Do While rdr.Read
                unitlist.Items.Add(rdr.Item("unit"))
            Loop
            rdr.Close()

        Catch sqlex As System.Data.SqlServerCe.SqlCeException
            MsgBox("Error reading prisms table in " & options(1) & vbNewLine & sqlex.Message, MsgBoxStyle.Exclamation, "Issue")
            Exit Try

        Finally
            units.Close()

        End Try

        minx.Text = ""
        miny.Text = ""
        maxx.Text = ""
        maxy.Text = ""
        radius.Text = ""
        centerx.Text = ""
        centery.Text = ""
        unitlist.SelectedItem = -1

    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click

        Dim u As String = unitlist.SelectedItem.ToString
        If u <> "" Then
            If MsgBox("Delete " + u + "?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "EDM-Mobile") = MsgBoxResult.Yes Then
                Cursor.Current = Cursors.WaitCursor
                Dim units As New SqlCeConnection("datasource='" + options(1) + "'")
                Dim sql As SqlCeCommand = units.CreateCommand
                units.Open()
                sql.CommandText = "DELETE FROM edm_units WHERE unit='" + u + "'"
                sql.ExecuteNonQuery()
                units.Close()
            Else
                Exit Sub
            End If
        End If
        populate_list()
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click

        Dim dsquare As Boolean = True
        Dim dcircle As Boolean = True

        If unitname.Text = "" Then
            MsgBox("Enter a name for the unit.", vbInformation, "EDMCE")
            Exit Sub
        End If

        If Not IsNumeric(minx.Text) Or Not IsNumeric(miny.Text) Or Not IsNumeric(maxx.Text) Or Not IsNumeric(maxy.Text) Then
            dsquare = False
        ElseIf CDbl(minx.Text) > CDbl(maxx.Text) Or CDbl(miny.Text) > CDbl(maxy.Text) Then
            MsgBox("Enter a maximum that is greater than the minimum.", vbInformation, "EDMCE")
            Exit Sub
        End If

        If Not IsNumeric(centerx.Text) Or Not IsNumeric(centery.Text) Or Not IsNumeric(radius.Text) Then
            dcircle = False
        End If

        If Not dcircle And Not dsquare Then
            MsgBox("Enter the coordinates for either a rectangular or a cirlce unit.", vbInformation, "EDMCE")
            Exit Sub
        End If

        Dim u As String = unitname.Text
        Dim tminx As String = minx.Text
        Dim tminy As String = miny.Text
        Dim tmaxx As String = maxx.Text
        Dim tmaxy As String = maxy.Text
        Dim tradius As String = radius.Text
        Dim tcenterx As String = centerx.Text
        Dim tcentery As String = centery.Text

        If tcentery = "" Then tcentery = "0.0"
        If tcenterx = "" Then tcenterx = "0.0"
        If tradius = "" Then tradius = "0.0"
        If tminx = "" Then tminx = "0.0"
        If tmaxx = "" Then tmaxx = "0.0"
        If tminy = "" Then tminy = "0.0"
        If tmaxy = "" Then tmaxy = "0.0"

        'check to see if this unit exists already
        Dim units As New SqlCeConnection("datasource =" + options(1))
        Dim sql As SqlCeCommand = units.CreateCommand
        sql.CommandText = "SELECT unit FROM edm_units WHERE unit='" + u + "'"
        units.Open()

        Dim r As Object = sql.ExecuteScalar()
        If LCase(CStr(r)) = LCase(u) Then
            If MsgBox("Overwrite existing unit?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "EDM-Mobile") = MsgBoxResult.No Then
                units.Close()
                Exit Sub
            End If
            Dim sqldel As SqlCeCommand = units.CreateCommand
            sqldel.CommandText = "DELETE FROM edm_units WHERE unit='" + u + "'"
            sqldel.ExecuteNonQuery()
        End If

        sql.CommandText = "INSERT INTO edm_units (unit,id,minx,miny,maxx,maxy,centerx,centery,radius) VALUES ('" + u + "',''," + tminx + "," + tminy + "," + tmaxx + "," + tmaxy + "," + tcenterx + "," + tcentery + "," + tradius + ")"
        sql.ExecuteNonQuery()
        units.Close()
        populate_list()

    End Sub

    Private Sub unitlist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles unitlist.SelectedIndexChanged

        If unitlist.SelectedIndex <> -1 Then
            Dim u As String = unitlist.SelectedItem.ToString
            Dim units As New SqlCeConnection("datasource='" + options(1) + "'")
            Dim sql As SqlCeCommand = units.CreateCommand
            Cursor.Current = Cursors.WaitCursor
            sql.CommandText = "SELECT unit,minx,miny,maxx,maxy,centerx,centery,radius FROM edm_units WHERE unit='" + u + "'"
            units.Open()
            Dim rdr As SqlCeDataReader = sql.ExecuteReader
            Do While rdr.Read
                If Not IsDBNull(rdr.Item("minx")) Then minx.Text = Format(rdr.Item("minx"), "##0.000") Else minx.Text = ""
                If Not IsDBNull(rdr.Item("minx")) Then miny.Text = Format(rdr.Item("miny"), "##0.000") Else miny.Text = ""
                If Not IsDBNull(rdr.Item("minx")) Then maxx.Text = Format(rdr.Item("maxx"), "##0.000") Else maxx.Text = ""
                If Not IsDBNull(rdr.Item("minx")) Then maxy.Text = Format(rdr.Item("maxy"), "##0.000") Else maxy.Text = ""
                If Not IsDBNull(rdr.Item("centerx")) Then centerx.Text = Format(rdr.Item("centerx"), "##0.000") Else centerx.Text = ""
                If Not IsDBNull(rdr.Item("centery")) Then centery.Text = Format(rdr.Item("centery"), "##0.000") Else centery.Text = ""
                If Not IsDBNull(rdr.Item("radius")) Then radius.Text = Format(rdr.Item("radius"), "##0.000") Else radius.Text = ""
                unitname.Text = rdr.Item("unit").ToString
            Loop
            units.Close()
            Cursor.Current = Cursors.Default
        End If

    End Sub

    Private Sub frmUnits_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Cursor.Current = Cursors.WaitCursor
        populate_list()

        Dim units As New SqlCeConnection("datasource='" + options(1) + "'")
        Dim sql As New SqlCeCommand("SELECT * FROM edm_units", units)
        Dim dataadapter As SqlCeDataAdapter = New SqlCeDataAdapter(sql)
        Dim ds As New DataSet
        units.Open()
        dataadapter.Fill(ds, "units")
        bs1.DataSource = ds
        bs1.DataMember = "units"
        unitsgrid.DataSource = bs1
        units.Close()

        Dim a As Integer
        fieldlist.BeginUpdate()
        For a = 1 To vars
            Select Case LCase(varlist(a))
                Case "x", "y", "z", "datumname", "datumx", "datumy", "datumz", "day", "date", "time", "unit", "suffix", "id", "hangle", "vangle", "prism"
                Case Else
                    fieldlist.Items.Add(varlist(a))
            End Select
        Next
        fieldlist.EndUpdate()

        unitfieldlist.BeginUpdate()
        Dim unitfields(100) As String
        Dim unitcount As Integer
        Call parse_list(options(17), ",", unitcount, unitfields)
        For a = 1 To unitcount
            Select Case LCase(unitfields(a))
                Case "id", "unit"
                Case Else
                    unitfieldlist.Items.Add(unitfields(a))
            End Select
        Next
        unitfieldlist.EndUpdate()

        Cursor.Current = Cursors.Default

    End Sub

    Private Sub fieldlist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles fieldlist.SelectedIndexChanged

        Dim flag As Boolean = False
        Dim a As Integer

        For a = 0 To unitfieldlist.Items.Count - 1
            If unitfieldlist.Items(a).ToString = fieldlist.Text Then
                flag = True
                Exit For
            End If
        Next
        If Not flag Then unitfieldlist.Items.Add(fieldlist.Text)

    End Sub

    Private Sub unitfieldlist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles unitfieldlist.SelectedIndexChanged

        If unitfieldlist.SelectedIndex = -1 Then Exit Sub
        unitfieldlist.Items.Remove(unitfieldlist.SelectedItem)

    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click

        Dim a As Integer
        unitfieldlist.BeginUpdate()
        unitfieldlist.Items.Clear()
        For a = 0 To fieldlist.Items.Count - 1
            unitfieldlist.Items.Add(fieldlist.Items(a))
        Next
        unitfieldlist.EndUpdate()

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click

        unitfieldlist.BeginUpdate()
        unitfieldlist.Items.Clear()
        unitfieldlist.EndUpdate()

    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        If unitfieldlist.SelectedIndex = -1 Then Exit Sub
        unitfieldlist.BeginUpdate()
        unitfieldlist.Items.Remove(unitfieldlist.SelectedItem)
        unitfieldlist.EndUpdate()

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub unitsgrid_CurrentCellChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles unitsgrid.CurrentCellChanged

    End Sub

    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged

        Select Case TabControl1.SelectedIndex
            Case 0
                MenuItem1.Enabled = True
                MenuItem3.Enabled = True
            Case 1, 2
                MenuItem1.Enabled = False
                MenuItem3.Enabled = False
            Case Else
        End Select
        Select Case TabControl1.SelectedIndex
            Case 2
                txtunitfields.Text = options(17)
            Case Else
        End Select

    End Sub

    Private Sub Label10_ParentChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label10.ParentChanged

    End Sub
End Class
