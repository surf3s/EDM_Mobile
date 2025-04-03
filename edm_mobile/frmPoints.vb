Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Data.SqlServerCe
Imports System.Windows.Forms

Public Class frmPoints
    Inherits System.Windows.Forms.Form
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents bs1 As System.Windows.Forms.BindingSource
    Private components As System.ComponentModel.IContainer
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem12 As System.Windows.Forms.MenuItem
    Friend WithEvents editcurrent As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem14 As System.Windows.Forms.MenuItem
    Friend WithEvents deletelastpoint As System.Windows.Forms.MenuItem
    Friend WithEvents deletecurrentpoint As System.Windows.Forms.MenuItem
    Friend WithEvents deletecurrentpointset As System.Windows.Forms.MenuItem
    Friend WithEvents deleteallpoints As System.Windows.Forms.MenuItem
    Friend WithEvents pointreshoot As System.Windows.Forms.MenuItem

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
    Friend WithEvents pointgrid As System.Windows.Forms.DataGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.pointreshoot = New System.Windows.Forms.MenuItem
        Me.editcurrent = New System.Windows.Forms.MenuItem
        Me.MenuItem14 = New System.Windows.Forms.MenuItem
        Me.deletelastpoint = New System.Windows.Forms.MenuItem
        Me.deletecurrentpoint = New System.Windows.Forms.MenuItem
        Me.deletecurrentpointset = New System.Windows.Forms.MenuItem
        Me.deleteallpoints = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.MenuItem9 = New System.Windows.Forms.MenuItem
        Me.MenuItem10 = New System.Windows.Forms.MenuItem
        Me.MenuItem11 = New System.Windows.Forms.MenuItem
        Me.MenuItem12 = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.bs1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.pointgrid = New System.Windows.Forms.DataGrid
        CType(Me.bs1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.Add(Me.MenuItem1)
        Me.MainMenu1.MenuItems.Add(Me.MenuItem2)
        Me.MainMenu1.MenuItems.Add(Me.MenuItem5)
        Me.MainMenu1.MenuItems.Add(Me.MenuItem6)
        '
        'MenuItem1
        '
        Me.MenuItem1.MenuItems.Add(Me.pointreshoot)
        Me.MenuItem1.MenuItems.Add(Me.editcurrent)
        Me.MenuItem1.MenuItems.Add(Me.MenuItem14)
        Me.MenuItem1.MenuItems.Add(Me.deletelastpoint)
        Me.MenuItem1.MenuItems.Add(Me.deletecurrentpoint)
        Me.MenuItem1.MenuItems.Add(Me.deletecurrentpointset)
        Me.MenuItem1.MenuItems.Add(Me.deleteallpoints)
        Me.MenuItem1.Text = "Points"
        '
        'pointreshoot
        '
        Me.pointreshoot.Text = "New XYZ for Current Point"
        '
        'editcurrent
        '
        Me.editcurrent.Text = "Edit Current Point"
        '
        'MenuItem14
        '
        Me.MenuItem14.Text = "-"
        '
        'deletelastpoint
        '
        Me.deletelastpoint.Text = "Delete Last Point"
        '
        'deletecurrentpoint
        '
        Me.deletecurrentpoint.Text = "Delete Current Point"
        '
        'deletecurrentpointset
        '
        Me.deletecurrentpointset.Text = "Delete Current Point Set"
        '
        'deleteallpoints
        '
        Me.deleteallpoints.Text = "Delete All Points"
        '
        'MenuItem2
        '
        Me.MenuItem2.MenuItems.Add(Me.MenuItem3)
        Me.MenuItem2.MenuItems.Add(Me.MenuItem4)
        Me.MenuItem2.Text = "Go to"
        '
        'MenuItem3
        '
        Me.MenuItem3.Text = "Top"
        '
        'MenuItem4
        '
        Me.MenuItem4.Text = "Bottom"
        '
        'MenuItem5
        '
        Me.MenuItem5.MenuItems.Add(Me.MenuItem7)
        Me.MenuItem5.MenuItems.Add(Me.MenuItem8)
        Me.MenuItem5.MenuItems.Add(Me.MenuItem9)
        Me.MenuItem5.MenuItems.Add(Me.MenuItem10)
        Me.MenuItem5.MenuItems.Add(Me.MenuItem11)
        Me.MenuItem5.MenuItems.Add(Me.MenuItem12)
        Me.MenuItem5.Text = "Sort"
        '
        'MenuItem7
        '
        Me.MenuItem7.Text = "Ascending"
        '
        'MenuItem8
        '
        Me.MenuItem8.Text = "Descending"
        '
        'MenuItem9
        '
        Me.MenuItem9.Text = "-"
        '
        'MenuItem10
        '
        Me.MenuItem10.Text = "By Unit-ID"
        '
        'MenuItem11
        '
        Me.MenuItem11.Text = "By Record No."
        '
        'MenuItem12
        '
        Me.MenuItem12.Text = "By Date and Time"
        '
        'MenuItem6
        '
        Me.MenuItem6.Text = "Close"
        '
        'pointgrid
        '
        Me.pointgrid.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.pointgrid.DataSource = Me.bs1
        Me.pointgrid.Location = New System.Drawing.Point(0, 1)
        Me.pointgrid.Name = "pointgrid"
        Me.pointgrid.Size = New System.Drawing.Size(240, 264)
        Me.pointgrid.TabIndex = 0
        '
        'frmPoints
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(240, 268)
        Me.ControlBox = False
        Me.Controls.Add(Me.pointgrid)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmPoints"
        Me.Text = "Points"
        CType(Me.bs1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub pointreshoot_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles pointreshoot.Click

        'First, get the recno of the current record
        Dim recno As Integer = CInt(pointgrid.Item(pointgrid.CurrentRowIndex, 0))

        'Second, get the point
        Dim whichprism As New frmWhichPrism
        shottype = 1
        If whichprism.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then

            'Third, update the record with the new XYZ
            Dim points As New SqlCeConnection("Datasource=" + options(1))
            Dim sql As SqlCeCommand = points.CreateCommand
            points.Open()
            sql.CommandText = "UPDATE [" & options(22) & "] SET X=" & currentxp & ",Y=" & currentyp & ",Z=" & currentzp & "WHERE recno=" & recno
            sql.ExecuteNonQuery()
            points.Close()
            Call loadpointstable()
        End If
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub MenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem6.Click

        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click

        Cursor.Current = Cursors.WaitCursor
        bs1.MoveFirst()
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click

        Cursor.Current = Cursors.WaitCursor
        bs1.MoveLast()
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub deleteallpoints_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles deleteallpoints.Click

        If MsgBox("Delete all records from table?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "EDM-Mobile") = MsgBoxResult.No Then Exit Sub

        If MsgBox("Once ALL records are deleted they cannot be recovered.  Proceed with delete?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "EDM-Mobile") = MsgBoxResult.No Then Exit Sub

        Cursor.Current = Cursors.WaitCursor

        Dim db As New SqlCeConnection("datasource='" + options(1) + "'")
        Dim sql As SqlCeCommand = db.CreateCommand
        sql.CommandText = "DELETE FROM [" & options(22) & "]"
        db.Open()
        sql.ExecuteNonQuery()
        db.Close()
        'pointgrid.Refresh()
        Call loadpointstable()
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub deletecurrentpointset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles deletecurrentpointset.Click

        'First, get the current unit-id
        Dim unit As String = ""
        Dim id As String = ""
        Dim suffix As String = ""
        Dim recno As Integer = 0
        Dim points As New SqlCeConnection("datasource='" + options(1) + "'")
        Dim sql As SqlCeCommand = points.CreateCommand
        sql.CommandText = "SELECT * FROM [" + options(22) + "]"
        points.Open()

        recno = CInt(DirectCast(DirectCast(bs1.Item(0), System.Object), System.Data.DataRowView).DataView.Item(pointgrid.CurrentRowIndex).Item(0))
        sql.CommandText = "SELECT unit,id,suffix,recno FROM [" + options(22) + "] WHERE recno=" & recno
        Dim rdr As SqlCeDataReader = sql.ExecuteReader
        If rdr.Read Then
            unit = rdr.Item(0).ToString
            id = rdr.Item(1).ToString
            suffix = rdr.Item(2).ToString
        End If
        rdr.Close()

        'Second, confirm deletion
        If MsgBox("Delete all records associated with " + Trim(unit) + "-" + Trim(id) + "?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.No Then
            points.Close()
            Exit Sub
        End If

        'Third, set-up query to delete all those records
        Cursor.Current = Cursors.WaitCursor
        'Dim points As New SqlCeConnection("datasource='" + options(1) + "'")
        Dim sqldel As SqlCeCommand = points.CreateCommand
        'points.Open()
        sqldel.CommandText = "DELETE FROM [" & options(22) & "] WHERE unit='" + unit + "' AND id='" + id + "'"
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

        points.Close()
        Call loadpointstable()
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub frmPoints_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Cursor.Current = Cursors.WaitCursor
        Call loadpointstable()
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub MenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem7.Click

    End Sub

    Private Sub editcurrent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles editcurrent.Click

        'First, fill the responses array with the current record
        Dim points As New SqlCeConnection("datasource='" + options(1) + "'")
        Dim sql As SqlCeCommand = points.CreateCommand
        sql.CommandText = "SELECT * FROM [" + options(22) + "]"
        points.Open()
        Dim rdr As SqlCeDataReader = sql.ExecuteReader

        Dim a As Integer
        Dim b As Integer
        For a = 0 To rdr.FieldCount - 1
            For b = 1 To vars
                If LCase(rdr.GetName(a)) = LCase(varlist(b)) Then
                    responses(b) = pointgrid.Item(pointgrid.CurrentRowIndex, a).ToString
                End If
            Next b
        Next a

        Dim rowno As Integer = pointgrid.CurrentRowIndex

        'Second, note the recno
        Dim recno As Integer = CInt(pointgrid.Item(pointgrid.CurrentRowIndex, 0))
        rdr.Close()

        'Third, show the edit form with these data
        Dim editpoint As New frmEditPoint
        back_to_edit = True
        If editpoint.ShowDialog() = Windows.Forms.DialogResult.OK Then

            'Fourth, put these new responses back into the record
            Dim id As String = ""
            Dim unit As String = ""
            'Dim points As New SqlCeConnection("Datasource=" + options(1))
            'Dim sql As SqlCeCommand = points.CreateCommand
            'points.Open()
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
            points.Close()

            'Dim flag As Boolean = False
            'bs1.ResetCurrentItem()
            'bs1.ResetBindings(False)
            'pointgrid.Update()
            'pointgrid.Refresh()
            Call loadpointstable()
            pointgrid.CurrentRowIndex = rowno
            pointgrid.Select(rowno)

        End If
        back_to_edit = False

    End Sub

    
    Private Sub deletecurrentpoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles deletecurrentpoint.Click

        'First, get the current unit, id, suffix and recno
        Dim unit As String = ""
        Dim id As String = ""
        Dim suffix As String = ""
        Dim recno As Integer = 0
        Dim rowno As Integer = pointgrid.CurrentRowIndex
        Dim points As New SqlCeConnection("datasource='" + options(1) + "'")
        Dim sql As SqlCeCommand = points.CreateCommand
        points.Open()

        recno = CInt(DirectCast(DirectCast(bs1.Item(0), System.Object), System.Data.DataRowView).DataView.Item(pointgrid.CurrentRowIndex).Item(0))
        sql.CommandText = "SELECT unit,id,suffix,recno FROM [" + options(22) + "] WHERE recno=" & recno
        Dim rdr As SqlCeDataReader = sql.ExecuteReader
        If rdr.Read Then
            unit = rdr.Item(0).ToString
            id = rdr.Item(1).ToString
            suffix = rdr.Item(2).ToString
        End If
        rdr.Close()

        'Second, confirm deletion
        If MsgBox("Delete record " + Trim(unit) + "-" + Trim(id) + "(" + CStr(suffix) + ") ?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.No Then Exit Sub

        'Third, set-up query to delete all those records
        Cursor.Current = Cursors.WaitCursor
        'Dim points As New SqlCeConnection("datasource='" + options(1) + "'")
        Dim sqldel As SqlCeCommand = points.CreateCommand
        'points.Open()
        sqldel.CommandText = "DELETE FROM [" & options(22) & "] WHERE recno=" & recno
        sqldel.ExecuteNonQuery()

        'See if ID deleted is current ID in units file and if so back off one
        If IsNumeric(id) Then
            If CInt(suffix) = 0 Then
                Dim r As Object
                'Dim sql As SqlCeCommand = points.CreateCommand
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
        points.Close()
        Call loadpointstable()
        If rowno + 1 > bs1.Count Then
            If bs1.Count > 0 Then rowno = bs1.Count - 1 Else rowno = -1
        End If
        If rowno <> -1 Then pointgrid.CurrentRowIndex = rowno
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub deletelastpoint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles deletelastpoint.Click

        Dim flag As Boolean
        Dim points As New SqlCeConnection("datasource='" + options(1) + "'")
        Dim sql As SqlCeCommand = points.CreateCommand
        Cursor.Current = Cursors.WaitCursor
        sql.CommandText = "SELECT * FROM [" & options(22) & "] ORDER BY Recno Desc"
        points.Open()
        Dim rdr As SqlCeDataReader = sql.ExecuteReader
        If rdr.Read Then
            flag = True
            Dim unit As String = rdr.Item("unit").ToString
            Dim id As String = rdr.Item("ID").ToString
            Dim suffix As String = rdr.Item("ID").ToString
            Dim sqidsuffix As String = unit + "-" + id + "(" + suffix + ")"
            If MsgBox("Delete record " + sqidsuffix + " ?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question) = MsgBoxResult.Yes Then
                Dim sqldel As SqlCeCommand = points.CreateCommand
                sqldel.CommandText = "DELETE FROM [" & options(22) & "] WHERE recno=" + rdr.Item("recno").ToString
                sqldel.ExecuteNonQuery()

                'See if ID deleted is current ID in units file and if so back off one
                If IsNumeric(id) Then
                    If CInt(suffix) = 0 Then
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
            End If

        Else
            MsgBox("The points table is empty.", MsgBoxStyle.Information)
            flag = False
        End If
        points.Close()
        If flag Then
            Call loadpointstable()
            bs1.MoveLast()
        End If
        Cursor.Current = Cursors.Default

    End Sub
    Private Sub loadpointstable()

        Dim points As New SqlCeConnection("datasource='" + options(1) + "'")
        Dim sql As New SqlCeCommand("SELECT * FROM " + options(22), points)
        Dim dataadapter As SqlCeDataAdapter = New SqlCeDataAdapter(sql)
        Dim ds As New DataSet
        points.Open()
        dataadapter.Fill(ds, options(22))
        bs1.DataSource = ds
        bs1.DataMember = options(22)
        pointgrid.DataSource = bs1
        points.Close()

    End Sub
End Class
