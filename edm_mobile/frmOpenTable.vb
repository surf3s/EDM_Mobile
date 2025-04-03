Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlServerCe
Imports System.IO
Imports System

Public Class frmOpenTable
    Inherits System.Windows.Forms.Form
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents tablelist As System.Windows.Forms.ListBox

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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.tablelist = New System.Windows.Forms.ListBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 32)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(224, 34)
        Me.Label1.Text = "Select from one of the following points files in this site database."
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(224, 16)
        Me.Label2.Text = "You are currently using"
        '
        'tablelist
        '
        Me.tablelist.Location = New System.Drawing.Point(8, 66)
        Me.tablelist.Name = "tablelist"
        Me.tablelist.Size = New System.Drawing.Size(224, 142)
        Me.tablelist.TabIndex = 2
        '
        'Button1
        '
        Me.Button1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular)
        Me.Button1.Location = New System.Drawing.Point(32, 218)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(176, 38)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Open Table"
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.Add(Me.MenuItem1)
        '
        'MenuItem1
        '
        Me.MenuItem1.Text = "Cancel"
        '
        'frmOpenTable
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(240, 268)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.tablelist)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmOpenTable"
        Me.Text = "Open Points Table"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Hide()
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click

        Me.DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub

    Private Sub frmOpenTable_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Cursor.Current = Cursors.WaitCursor
        tablelist.BeginUpdate()
        Dim systemtables As New SqlCeConnection("Datasource='" + options(1) + "'")
        Dim sql As SqlCeCommand = systemtables.CreateCommand
        sql.CommandText = "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES"
        Dim rdr As SqlCeDataReader
        systemtables.Open()
        rdr = sql.ExecuteReader
        Do While rdr.Read
            Select LCase(rdr.Item("TABLE_NAME").ToString)
                Case LCase(options(22))
                Case "edm_units"
                Case "edm_prisms"
                Case "edm_datums"
                Case Else
                    tablelist.Items.Add(rdr.Item("TABLE_NAME"))
            End Select
        Loop
        rdr.Close()
        systemtables.Close()
        tablelist.EndUpdate()
        Label2.Text = "Currently using " + options(22)
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim status As Integer

        If tablelist.SelectedIndex = -1 Then
            MsgBox("Select a point tables from the list.", vbOKOnly, "EDM-Mobile")
            Exit Sub
        End If

        If tablelist.SelectedItem.ToString <> options(22) Then
            Cursor.Current = Cursors.WaitCursor
            Dim iniclass As String = "[EDM]"
            Dim inidata(1, 2) As String

            options(22) = tablelist.SelectedItem.ToString
            inidata(1, 1) = "POINTTABLE"
            inidata(1, 2) = options(22)
            Call WriteIni(cfgfile, iniclass, inidata, status)
            Cursor.Current = Cursors.Default
        End If
        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub
End Class
