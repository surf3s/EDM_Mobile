Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlServerCe
Imports System.io
Imports System

Public Class frmPrisms
    Inherits System.Windows.Forms.Form
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents prismlist As System.Windows.Forms.ListBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents prismname As System.Windows.Forms.TextBox
    Friend WithEvents prismheight As System.Windows.Forms.TextBox
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
    Friend WithEvents prismconstant As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPrisms))
        Me.Label1 = New System.Windows.Forms.Label
        Me.prismlist = New System.Windows.Forms.ListBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.prismname = New System.Windows.Forms.TextBox
        Me.prismheight = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.prismconstant = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(142, 132)
        Me.Label1.Text = resources.GetString("Label1.Text")
        '
        'prismlist
        '
        Me.prismlist.Location = New System.Drawing.Point(156, 8)
        Me.prismlist.Name = "prismlist"
        Me.prismlist.Size = New System.Drawing.Size(76, 142)
        Me.prismlist.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 154)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 16)
        Me.Label2.Text = "Name:"
        '
        'prismname
        '
        Me.prismname.Location = New System.Drawing.Point(8, 170)
        Me.prismname.MaxLength = 10
        Me.prismname.Name = "prismname"
        Me.prismname.Size = New System.Drawing.Size(80, 21)
        Me.prismname.TabIndex = 0
        '
        'prismheight
        '
        Me.prismheight.Location = New System.Drawing.Point(104, 170)
        Me.prismheight.MaxLength = 15
        Me.prismheight.Name = "prismheight"
        Me.prismheight.Size = New System.Drawing.Size(56, 21)
        Me.prismheight.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(104, 154)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(40, 16)
        Me.Label3.Text = "Height:"
        '
        'prismconstant
        '
        Me.prismconstant.Location = New System.Drawing.Point(176, 170)
        Me.prismconstant.MaxLength = 15
        Me.prismconstant.Name = "prismconstant"
        Me.prismconstant.Size = New System.Drawing.Size(56, 21)
        Me.prismconstant.TabIndex = 2
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(176, 154)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(56, 16)
        Me.Label4.Text = "Constant:"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(8, 197)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(224, 56)
        Me.Label5.Text = "If the prism constants (also called the offset) are managed incorrectly it will r" & _
            "esult in systematic and difficult to correct errors."
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.Add(Me.MenuItem1)
        Me.MainMenu1.MenuItems.Add(Me.MenuItem2)
        Me.MainMenu1.MenuItems.Add(Me.MenuItem3)
        '
        'MenuItem1
        '
        Me.MenuItem1.Text = "Save          "
        '
        'MenuItem2
        '
        Me.MenuItem2.Text = "Delete        "
        '
        'MenuItem3
        '
        Me.MenuItem3.Text = "Close"
        '
        'frmPrisms
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(240, 268)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.prismconstant)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.prismheight)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.prismname)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.prismlist)
        Me.Controls.Add(Me.Label1)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmPrisms"
        Me.Text = "Prisms"
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub prismlist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles prismlist.SelectedIndexChanged

        If prismlist.SelectedIndex <> -1 Then
            Dim p As String = prismlist.SelectedItem.ToString
            Dim prisms As New SqlCeConnection("datasource='" + options(1) + "'")
            Dim sql As SqlCeCommand = prisms.CreateCommand
            Cursor.Current = Cursors.WaitCursor
            sql.CommandText = "SELECT name,height,offset FROM edm_poles WHERE name='" + p + "'"
            prisms.Open()
            Dim rdr As SqlCeDataReader = sql.ExecuteReader
            Do While rdr.Read
                prismheight.Text = Format(rdr.Item("height"), "##0.000")
                prismconstant.Text = Format(rdr.Item("offset"), "0.000")
                prismname.Text = rdr.Item("name").ToString
            Loop
            prisms.Close()
            Cursor.Current = Cursors.Default
        End If

    End Sub

    Private Sub prismlist_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles prismlist.Click
    End Sub

    Private Sub frmPrisms_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        populate_list()

    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click

        If prismname.Text = "" Then
            MsgBox("Enter a name for the prism", MsgBoxStyle.Information, "EDM-Mobile")
            Exit Sub
        End If

        If Not IsNumeric(prismheight.Text) Then
            MsgBox("Enter a prism height (meters or feet) for this prism.", MsgBoxStyle.Information, "EDM-Mobile")
            Exit Sub
        End If

        If Not IsNumeric(prismconstant.Text) Then
            MsgBox("Enter a valid prism offset (meters or feet) or none at all.  See ? on this form for more information about offsets", MsgBoxStyle.Information, "EDM-Mobile")
            Exit Sub
        End If

        'check to see if this prism exists already
        Dim p As String = prismname.Text
        Dim pheight As String = prismheight.Text
        Dim poffset As String = prismconstant.Text
        Dim prisms As New SqlCeConnection("datasource =" + options(1))
        Dim sql As SqlCeCommand = prisms.CreateCommand
        sql.CommandText = "SELECT name FROM edm_poles WHERE name='" + p + "'"
        prisms.Open()

        Dim r As Object = sql.ExecuteScalar()
        If LCase(CStr(r)) = LCase(p) Then
            If MsgBox("Overwrite existing prism?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "EDM-Mobile") = MsgBoxResult.No Then
                prisms.Close()
                Exit Sub
            End If
            Dim sqldel As SqlCeCommand = prisms.CreateCommand
            sqldel.CommandText = "DELETE FROM edm_poles WHERE name='" + p + "'"
            sqldel.ExecuteNonQuery()
        End If

        If prismconstant.Text = "" Then prismconstant.Text = "0.000"

        sql.CommandText = "INSERT INTO edm_poles (name,height,offset) VALUES ('" + p + "'," + pheight + "," + poffset + ")"
        sql.ExecuteNonQuery()
        prisms.Close()
        populate_list()

    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        If prismlist.SelectedIndex = -1 Then Exit Sub
        Dim p As String = prismlist.SelectedItem.ToString
        If p <> "" Then
            If MsgBox("Delete " + p + "?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "EDM-Mobile") = MsgBoxResult.Yes Then
                Cursor.Current = Cursors.WaitCursor
                Dim prisms As New SqlCeConnection("datasource='" + options(1) + "'")
                Dim sql As SqlCeCommand = prisms.CreateCommand
                prisms.Open()
                sql.CommandText = "DELETE FROM edm_poles WHERE name='" + p + "'"
                sql.ExecuteNonQuery()
                prisms.Close()
            Else
                Exit Sub
            End If
        End If
        populate_list()
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click

        npoles = 0    'force a reread of array that stores these
        Me.DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub
    Private Sub populate_list()

        Dim prisms As New SqlCeConnection("datasource='" + options(1) + "'")
        Dim sql As SqlCeCommand = prisms.CreateCommand

        Cursor.Current = Cursors.WaitCursor

        Try
            prisms.Open()
            sql.CommandText = "SELECT * FROM edm_poles"
            Dim rdr As SqlCeDataReader = sql.ExecuteReader
            prismlist.Items.Clear()
            Do While rdr.Read
                prismlist.Items.Add(rdr.Item("name"))
            Loop
            rdr.Close()

        Catch sqlex As System.Data.SqlServerCe.SqlCeException
            MsgBox("Error reading prisms table in " & options(1) & vbNewLine & sqlex.Message, MsgBoxStyle.Exclamation, "Issue")
            Exit Try

        Finally
            prisms.Close()

        End Try

        prismheight.Text = ""
        prismconstant.Text = ""
        prismname.Text = ""
        prismlist.SelectedItem = -1
        Cursor.Current = Cursors.Default

    End Sub
End Class
