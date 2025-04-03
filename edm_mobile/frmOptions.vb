Public Class frmOptions
    Inherits System.Windows.Forms.Form
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents prismoffsetcheck As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents autolocate As System.Windows.Forms.CheckBox
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents barcodecheck As System.Windows.Forms.CheckBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents writelogfilecheck As System.Windows.Forms.CheckBox
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem

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
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.TabPage2 = New System.Windows.Forms.TabPage
        Me.Label2 = New System.Windows.Forms.Label
        Me.prismoffsetcheck = New System.Windows.Forms.CheckBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.autolocate = New System.Windows.Forms.CheckBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.writelogfilecheck = New System.Windows.Forms.CheckBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.barcodecheck = New System.Windows.Forms.CheckBox
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.Add(Me.MenuItem1)
        '
        'MenuItem1
        '
        Me.MenuItem1.Text = "Close"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(240, 267)
        Me.TabControl1.TabIndex = 7
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.Label2)
        Me.TabPage1.Controls.Add(Me.prismoffsetcheck)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Controls.Add(Me.autolocate)
        Me.TabPage1.Location = New System.Drawing.Point(0, 0)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(240, 244)
        Me.TabPage1.Text = "Options 1"
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.Label4)
        Me.TabPage2.Controls.Add(Me.barcodecheck)
        Me.TabPage2.Controls.Add(Me.Label3)
        Me.TabPage2.Controls.Add(Me.writelogfilecheck)
        Me.TabPage2.Location = New System.Drawing.Point(0, 0)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Size = New System.Drawing.Size(240, 244)
        Me.TabPage2.Text = "Options 2"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(18, 142)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(204, 77)
        Me.Label2.Text = "This options lets you hide prism offset menus to prevent errors.  Use this option" & _
            " if prism offsets are set in the total station and should not be changed."
        '
        'prismoffsetcheck
        '
        Me.prismoffsetcheck.Location = New System.Drawing.Point(18, 120)
        Me.prismoffsetcheck.Name = "prismoffsetcheck"
        Me.prismoffsetcheck.Size = New System.Drawing.Size(204, 19)
        Me.prismoffsetcheck.TabIndex = 7
        Me.prismoffsetcheck.Text = "Allow changes to prism offsets?"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(18, 38)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(204, 77)
        Me.Label1.Text = "This options uses the XY coordinates of a point to automatically switch to the de" & _
            "fined Unit.  To use this option you need to setup Units and Unit Fields."
        '
        'autolocate
        '
        Me.autolocate.Location = New System.Drawing.Point(18, 16)
        Me.autolocate.Name = "autolocate"
        Me.autolocate.Size = New System.Drawing.Size(188, 19)
        Me.autolocate.TabIndex = 5
        Me.autolocate.Text = "Auto-locate units ?"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(18, 41)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(204, 31)
        Me.Label3.Text = "Writes station initialization data to a log file with same name as CFG file."
        '
        'writelogfilecheck
        '
        Me.writelogfilecheck.Location = New System.Drawing.Point(18, 19)
        Me.writelogfilecheck.Name = "writelogfilecheck"
        Me.writelogfilecheck.Size = New System.Drawing.Size(204, 19)
        Me.writelogfilecheck.TabIndex = 8
        Me.writelogfilecheck.Text = "Write Log file?"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(18, 118)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(204, 31)
        Me.Label4.Text = "This option launches a shot when a Unit-ID barcode is scanned."
        '
        'barcodecheck
        '
        Me.barcodecheck.Location = New System.Drawing.Point(18, 96)
        Me.barcodecheck.Name = "barcodecheck"
        Me.barcodecheck.Size = New System.Drawing.Size(204, 19)
        Me.barcodecheck.TabIndex = 10
        Me.barcodecheck.Text = "Barcode Pre-Scan?"
        '
        'frmOptions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(240, 270)
        Me.ControlBox = False
        Me.Controls.Add(Me.TabControl1)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmOptions"
        Me.Text = "EDM-Mobile Options"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click

        Dim status As Integer
        Dim iniclass As String
        Dim inidata(4, 2) As String

        iniclass = "[EDM]"
        inidata(1, 1) = "LimitChecking"
        If use_limitchecking Then
            inidata(1, 2) = "Yes"
        Else
            inidata(1, 2) = "No"
        End If

        inidata(2, 1) = "ShowOffsetMenu"
        If showoffsetmenu Then
            inidata(2, 2) = "Yes"
        Else
            inidata(2, 2) = "No"
        End If

        inidata(3, 1) = "WriteLogFile"
        If WriteLogFile Then
            inidata(3, 2) = "Yes"
        Else
            inidata(3, 2) = "No"
        End If

        inidata(4, 1) = "BarcodePreScan"
        If bar_code_pre_scan Then
            inidata(4, 2) = "Yes"
        Else
            inidata(4, 2) = "No"
        End If

        Call WriteIni(cfgfile, iniclass, inidata, status)

        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub

    Private Sub frmOptions_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If use_limitchecking Then
            autolocate.Checked = True
        Else
            autolocate.Checked = False
        End If

        If showoffsetmenu Then
            prismoffsetcheck.Checked = True
        Else
            prismoffsetcheck.Checked = False
        End If

        If WriteLogFile Then
            writelogfilecheck.Checked = True
        Else
            writelogfilecheck.Checked = False
        End If

        If bar_code_pre_scan Then
            barcodecheck.Checked = True
        Else
            barcodecheck.Checked = False
        End If

        Cursor.Current = Cursors.Default

    End Sub

    Private Sub autolocate_CheckStateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Select Case autolocate.CheckState
            Case CheckState.Checked
                use_limitchecking = True
            Case CheckState.Unchecked
                use_limitchecking = False
            Case Else
        End Select

    End Sub

    Private Sub prismoffsetcheck_CheckStateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Select Case prismoffsetcheck.CheckState
            Case CheckState.Checked
                showoffsetmenu = True
            Case CheckState.Unchecked
                showoffsetmenu = False
            Case Else
        End Select

    End Sub

    Private Sub writelogfilecheck_CheckStateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Select Case writelogfilecheck.CheckState
            Case CheckState.Checked
                WriteLogFile = True
            Case CheckState.Unchecked
                WriteLogFile = False
            Case Else
        End Select

    End Sub

    Private Sub writelogfilecheck_CheckStateChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles writelogfilecheck.CheckStateChanged

    End Sub

    Private Sub barcodecheck_CheckStateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles barcodecheck.CheckStateChanged

        Select Case barcodecheck.CheckState
            Case CheckState.Checked
                bar_code_pre_scan = True

            Case CheckState.Unchecked
                bar_code_pre_scan = False

            Case Else
        End Select

    End Sub
End Class
