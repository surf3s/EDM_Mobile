Imports System.Data.SqlClient
Imports System.Data.SqlServerCe
Public Class frmNewTable
    Inherits System.Windows.Forms.Form
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtnewtable As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button

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
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtnewtable = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(32, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(176, 48)
        Me.Label1.Text = "Use this option to start a new points file but keep the same datums, units and pr" & _
            "isms."
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(32, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(176, 56)
        Me.Label2.Text = "Enter the new points file name (this will become a table in the existing site dat" & _
            "abase) :"
        '
        'txtnewtable
        '
        Me.txtnewtable.Location = New System.Drawing.Point(32, 120)
        Me.txtnewtable.Name = "txtnewtable"
        Me.txtnewtable.Size = New System.Drawing.Size(176, 21)
        Me.txtnewtable.TabIndex = 1
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(80, 152)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(72, 24)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Create"
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.Add(Me.MenuItem1)
        '
        'MenuItem1
        '
        Me.MenuItem1.Text = "Close"
        '
        'frmNewTable
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(240, 270)
        Me.ControlBox = False
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.txtnewtable)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmNewTable"
        Me.Text = "EDM-Mobile"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click

        Me.Hide()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim status As Integer
        Dim tbname As String
        Dim dbname As String
        Dim iniclass As String
        Dim inidata(1, 2) As String
        Dim a As Integer

        For a = 1 To Len(txtnewtable.Text)
            If InStr("!@#$%^&*() {}[]<>,.?/|\`~" & Chr(34), Mid(txtnewtable.Text, a, 1)) <> 0 Then
                MsgBox("This name contains illegal characters.  Spaces are also not allowed.", vbInformation And MsgBoxStyle.OkOnly, "EDM-Mobile")
                Exit Sub
            End If
        Next a

        If LCase(txtnewtable.Text) = LCase(options(22)) Then
            MsgBox(options(22) + " is already the current table. Please supply a different name.  If you want to erase the data in " + options(22) + " use Edit-Points-Delete All.", MsgBoxStyle.Information And MsgBoxStyle.OkOnly, "EDM-Mobile")
            Exit Sub
        End If

        Cursor.Current = Cursors.WaitCursor
        tbname = txtnewtable.Text
        Call table_status(tbname, status)
        Cursor.Current = Cursors.Default

        If status > 0 Then
            'Table already exists and contains data
            If MsgBox("This points file already exists and contains " + CStr(status) + " records.  Do you want to overwrite it?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "EDM-Mobile") = MsgBoxResult.Yes Then
                If MsgBox("Sorry, but are you sure you want to overwrite the points file " + tbname + " ?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "EDM-Mobile") = MsgBoxResult.Yes Then
                    Cursor.Current = Cursors.WaitCursor
                    Dim points As New SqlCeConnection("Datasource='" + options(1) + "'")
                    Dim pointssql As SqlCeCommand = points.CreateCommand
                    pointssql.CommandText = "DELETE [" + tbname + "]"
                    Dim r As Object = pointssql.ExecuteNonQuery
                    points.Close()
                    dbname = options(1)
                    options(22) = tbname
                    Call add_field_tb()
                Else
                    Exit Sub
                End If
            Else
                Exit Sub
            End If

        ElseIf status = 0 Then
            'nothing to do - the table exists and is empty
            options(22) = tbname

        Else
            'table does not exists
            Cursor.Current = Cursors.WaitCursor
            options(22) = tbname
            Call add_field_tb()

        End If

        Cursor.Current = Cursors.WaitCursor
        iniclass = "[EDM]"
        inidata(1, 1) = "POINTTABLE"
        inidata(1, 2) = options(22)
        Call WriteIni(cfgfile, iniclass, inidata, status)
        Cursor.Current = Cursors.Default

        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub

    Private Sub frmNewTable_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Cursor.Current = Cursors.Default

    End Sub
End Class
