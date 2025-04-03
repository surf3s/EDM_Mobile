Public Class frmFieldSetup
    Inherits System.Windows.Forms.Form
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents fieldname As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents fieldtype As System.Windows.Forms.ComboBox
    Friend WithEvents prompt As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents fieldsize As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents menulist As System.Windows.Forms.ListBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label

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
    Friend WithEvents addmenu As System.Windows.Forms.Button
    Friend WithEvents deletemenu As System.Windows.Forms.Button
    Friend WithEvents txtaddmenu As System.Windows.Forms.TextBox
    Friend WithEvents txtrequired As System.Windows.Forms.CheckBox
    Friend WithEvents txtincrement As System.Windows.Forms.CheckBox
    Friend WithEvents txtcarry As System.Windows.Forms.CheckBox
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.txtcarry = New System.Windows.Forms.CheckBox
        Me.txtincrement = New System.Windows.Forms.CheckBox
        Me.txtrequired = New System.Windows.Forms.CheckBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.fieldname = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.fieldtype = New System.Windows.Forms.ComboBox
        Me.prompt = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.fieldsize = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.menulist = New System.Windows.Forms.ListBox
        Me.addmenu = New System.Windows.Forms.Button
        Me.deletemenu = New System.Windows.Forms.Button
        Me.txtaddmenu = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.Label7 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'txtcarry
        '
        Me.txtcarry.Location = New System.Drawing.Point(8, 184)
        Me.txtcarry.Name = "txtcarry"
        Me.txtcarry.Size = New System.Drawing.Size(112, 24)
        Me.txtcarry.TabIndex = 17
        Me.txtcarry.Text = "Carry last value"
        '
        'txtincrement
        '
        Me.txtincrement.Location = New System.Drawing.Point(8, 208)
        Me.txtincrement.Name = "txtincrement"
        Me.txtincrement.Size = New System.Drawing.Size(112, 24)
        Me.txtincrement.TabIndex = 16
        Me.txtincrement.Text = "Auto increment"
        '
        'txtrequired
        '
        Me.txtrequired.Location = New System.Drawing.Point(120, 184)
        Me.txtrequired.Name = "txtrequired"
        Me.txtrequired.Size = New System.Drawing.Size(112, 24)
        Me.txtrequired.TabIndex = 15
        Me.txtrequired.Text = "Required"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 16)
        Me.Label1.Text = "Field Name :"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'fieldname
        '
        Me.fieldname.Location = New System.Drawing.Point(96, 8)
        Me.fieldname.Name = "fieldname"
        Me.fieldname.Size = New System.Drawing.Size(136, 21)
        Me.fieldname.TabIndex = 12
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 32)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(80, 16)
        Me.Label2.Text = "Type :"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'fieldtype
        '
        Me.fieldtype.Items.Add("Text")
        Me.fieldtype.Items.Add("Numeric")
        Me.fieldtype.Items.Add("Instrument")
        Me.fieldtype.Items.Add("Menu")
        Me.fieldtype.Items.Add("Unit")
        Me.fieldtype.Location = New System.Drawing.Point(96, 32)
        Me.fieldtype.Name = "fieldtype"
        Me.fieldtype.Size = New System.Drawing.Size(136, 22)
        Me.fieldtype.TabIndex = 10
        '
        'prompt
        '
        Me.prompt.Location = New System.Drawing.Point(96, 56)
        Me.prompt.Name = "prompt"
        Me.prompt.Size = New System.Drawing.Size(136, 21)
        Me.prompt.TabIndex = 8
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 56)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(80, 16)
        Me.Label3.Text = "Prompt :"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'fieldsize
        '
        Me.fieldsize.Location = New System.Drawing.Point(96, 80)
        Me.fieldsize.Name = "fieldsize"
        Me.fieldsize.Size = New System.Drawing.Size(136, 21)
        Me.fieldsize.TabIndex = 6
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 80)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(80, 16)
        Me.Label4.Text = "Size :"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(0, 104)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(48, 16)
        Me.Label5.Text = "Menu :"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'menulist
        '
        Me.menulist.Location = New System.Drawing.Point(48, 104)
        Me.menulist.Name = "menulist"
        Me.menulist.Size = New System.Drawing.Size(184, 44)
        Me.menulist.TabIndex = 4
        '
        'addmenu
        '
        Me.addmenu.Location = New System.Drawing.Point(144, 152)
        Me.addmenu.Name = "addmenu"
        Me.addmenu.Size = New System.Drawing.Size(40, 24)
        Me.addmenu.TabIndex = 2
        Me.addmenu.Text = "Add"
        '
        'deletemenu
        '
        Me.deletemenu.Location = New System.Drawing.Point(192, 152)
        Me.deletemenu.Name = "deletemenu"
        Me.deletemenu.Size = New System.Drawing.Size(40, 24)
        Me.deletemenu.TabIndex = 3
        Me.deletemenu.Text = "Del"
        '
        'txtaddmenu
        '
        Me.txtaddmenu.Location = New System.Drawing.Point(72, 152)
        Me.txtaddmenu.Name = "txtaddmenu"
        Me.txtaddmenu.Size = New System.Drawing.Size(64, 21)
        Me.txtaddmenu.TabIndex = 0
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(0, 152)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(72, 16)
        Me.Label6.Text = "New Item :"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.Add(Me.MenuItem1)
        Me.MainMenu1.MenuItems.Add(Me.MenuItem2)
        Me.MainMenu1.MenuItems.Add(Me.MenuItem3)
        '
        'MenuItem1
        '
        Me.MenuItem1.Text = "Previous Field "
        '
        'MenuItem2
        '
        Me.MenuItem2.Text = "Next Field  "
        '
        'MenuItem3
        '
        Me.MenuItem3.Text = "Close"
        '
        'Label7
        '
        Me.Label7.Location = New System.Drawing.Point(40, 232)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(155, 30)
        Me.Label7.Text = "(To add fields to the Unit defaults see Edit-Units)"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'frmFieldSetup
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(240, 268)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtaddmenu)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.addmenu)
        Me.Controls.Add(Me.deletemenu)
        Me.Controls.Add(Me.menulist)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.fieldsize)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.prompt)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.fieldtype)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.fieldname)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtrequired)
        Me.Controls.Add(Me.txtincrement)
        Me.Controls.Add(Me.txtcarry)
        Me.Menu = Me.MainMenu1
        Me.Name = "frmFieldSetup"
        Me.Text = "Field Setup"
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim showvar As Integer
    Dim inprogress As Boolean
    Dim changes As Boolean

    Private Sub frmFieldSetup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        showvar = 1
        Call load_var(showvar)
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub frmFieldSetup_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated

    End Sub
    Private Sub load_var(ByVal v As Integer)

        Dim t1 As String
        Dim a As Integer
        Dim t2 As String

        inprogress = True
        fieldname.Text = varlist(v)
        prompt.Text = varprompt(v)
        fieldsize.Text = vlen(v).ToString
        menulist.Items.Clear()
        txtaddmenu.Enabled = False
        addmenu.Enabled = False
        deletemenu.Enabled = False
        Select Case vtype(v)
            Case 1
                fieldtype.Text = "Text"
            Case 2
                fieldtype.Text = "Numeric"
            Case 3
                fieldtype.Text = "Instrument"
            Case 4
                fieldtype.Text = "Menu"
                t1 = vmenu(v)
                Do Until t1 = ""
                    a = InStr(t1, ",")
                    If a <> 0 Then
                        t2 = leftstr(t1, a - 1)
                        If t2 <> "" Then menulist.Items.Add(t2)
                        t1 = Mid(t1, a + 1)
                    ElseIf t1 <> "" Then
                        menulist.Items.Add(t1)
                        Exit Do
                    End If
                Loop
                txtaddmenu.Enabled = True
                addmenu.Enabled = True
                deletemenu.Enabled = True
            Case 5
                fieldtype.Text = "Unit"
            Case Else
        End Select
        If varrequired(v) = True Then
            txtrequired.Checked = True
        Else
            txtrequired.Checked = False
        End If
        If vincr(v) = True Then
            txtincrement.Checked = True
        Else
            txtincrement.Checked = False
        End If
        If vcarry(v) = True Then
            txtcarry.Checked = True
        Else
            txtcarry.Checked = False
        End If

        changes = False
        inprogress = False

    End Sub

    Private Sub CheckBox1_CheckStateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtincrement.CheckStateChanged
        If inprogress Then Exit Sub
        changes = True
    End Sub

    Private Sub addmenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles addmenu.Click

        Dim flag As Boolean
        Dim a As Integer

        If txtaddmenu.Text <> "" Then
            flag = False
            For a = 0 To menulist.Items.Count - 1
                If LCase(txtaddmenu.Text) = menulist.Items(a).ToString Then
                    flag = True
                    Exit For
                End If
            Next a
            If Not flag Then
                menulist.Items.Add(txtaddmenu.Text)
                txtaddmenu.Text = ""
                changes = True
            End If
        End If

    End Sub

    Private Sub deletemenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles deletemenu.Click

        Dim answer As Integer
        If menulist.SelectedItem.ToString <> "" Then
            answer = MsgBox("Delete " + menulist.SelectedItem.ToString + "?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "EDM-Mobile")
            If answer = 6 Then
                menulist.Items.Remove(menulist.SelectedItem)
                changes = True
            End If
        End If
    End Sub

    Private Sub txtcarry_CheckStateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtcarry.CheckStateChanged
        If inprogress Then Exit Sub
        changes = True
    End Sub

    Private Sub txtrequired_CheckStateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtrequired.CheckStateChanged
        If inprogress Then Exit Sub
        changes = True
    End Sub

    Private Sub unitfield_CheckStateChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If inprogress Then Exit Sub
        changes = True
    End Sub

    Private Sub prompt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles prompt.TextChanged
        If inprogress Then Exit Sub
        changes = True
    End Sub
    Private Sub save_var()

        Dim iniclass As String
        Dim inidata(5, 2) As String
        Dim a As Integer
        Dim wstatus As Integer

        iniclass = "[" + varlist(showvar) + "]"
        inidata(1, 1) = "Prompt"
        inidata(1, 2) = prompt.Text
        varprompt(showvar) = prompt.Text
        inidata(2, 1) = "Carry"
        If Not txtcarry.Checked Then
            inidata(2, 2) = "No"
            vcarry(showvar) = False
        Else
            inidata(2, 2) = "Yes"
            vcarry(showvar) = True
        End If
        inidata(3, 1) = "Increment"
        If Not txtincrement.Checked Then
            inidata(3, 2) = "No"
            vincr(showvar) = False
        Else
            inidata(3, 2) = "Yes"
            vincr(showvar) = True
        End If
        If vtype(showvar) = 4 Then
            vmenu(showvar) = ""
            For a = 0 To menulist.Items.Count - 1
                If vmenu(showvar) = "" Then
                    vmenu(showvar) = menulist.Items(a).ToString
                Else
                    vmenu(showvar) = vmenu(showvar) & "," & menulist.Items(a).ToString
                End If
            Next a
            inidata(4, 1) = "Menu"
            inidata(4, 2) = vmenu(showvar)
        End If
        inidata(5, 1) = "Required"
        If Not txtrequired.Checked Then
            inidata(5, 2) = "No"
            varrequired(showvar) = False
        Else
            inidata(5, 2) = "Yes"
            varrequired(showvar) = True
        End If
        Call WriteIni(cfgfile, iniclass, inidata, wstatus)
        changes = False

    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        If showvar <> vars Then
            Cursor.Current = Cursors.WaitCursor
            If changes Then Call save_var()
            showvar = showvar + 1
            Call load_var(showvar)
            Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        If showvar <> 1 Then
            Cursor.Current = Cursors.WaitCursor
            If changes Then Call save_var()
            showvar = showvar - 1
            Call load_var(showvar)
            Cursor.Current = Cursors.Default
        End If
    End Sub

    Private Sub MenuItem3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuItem3.Click

        If changes Then Call save_var()
        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub
End Class
