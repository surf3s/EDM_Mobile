Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlServerCe
Imports System.io
Imports System
Public Class frmEditPoint
    Inherits System.Windows.Forms.Form
    Dim firstvar As Integer
    Dim inprocessedit As Boolean
    Dim temp_prisms(100) As Single
    Dim putback As Single
    Dim skip_var As Integer
    Dim varlabel(9) As Label
    Dim varcombo(9) As ComboBox
    Dim vartext(9) As TextBox
    Dim savebutton(9) As Button
    Friend WithEvents vartext10 As System.Windows.Forms.TextBox
    Friend WithEvents vartext11 As System.Windows.Forms.TextBox
    Friend WithEvents var10combo As System.Windows.Forms.ComboBox
    Friend WithEvents var11combo As System.Windows.Forms.ComboBox
    Friend WithEvents vno10 As System.Windows.Forms.Label
    Friend WithEvents vno11 As System.Windows.Forms.Label
    Friend WithEvents varlabel10 As System.Windows.Forms.Label
    Friend WithEvents varlabel11 As System.Windows.Forms.Label
    Dim vno(9) As Label

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
    Friend WithEvents VScroll As System.Windows.Forms.VScrollBar
    Friend WithEvents varlabel1 As System.Windows.Forms.Label
    Friend WithEvents vartext1 As System.Windows.Forms.TextBox
    Friend WithEvents var1combo As System.Windows.Forms.ComboBox
    Friend WithEvents vno1 As System.Windows.Forms.Label
    Friend WithEvents var2combo As System.Windows.Forms.ComboBox
    Friend WithEvents vno2 As System.Windows.Forms.Label
    Friend WithEvents vartext2 As System.Windows.Forms.TextBox
    Friend WithEvents varlabel2 As System.Windows.Forms.Label
    Friend WithEvents var3combo As System.Windows.Forms.ComboBox
    Friend WithEvents vno3 As System.Windows.Forms.Label
    Friend WithEvents vartext3 As System.Windows.Forms.TextBox
    Friend WithEvents varlabel3 As System.Windows.Forms.Label
    Friend WithEvents var4combo As System.Windows.Forms.ComboBox
    Friend WithEvents vno4 As System.Windows.Forms.Label
    Friend WithEvents vartext4 As System.Windows.Forms.TextBox
    Friend WithEvents varlabel4 As System.Windows.Forms.Label
    Friend WithEvents var5combo As System.Windows.Forms.ComboBox
    Friend WithEvents vno5 As System.Windows.Forms.Label
    Friend WithEvents vartext5 As System.Windows.Forms.TextBox
    Friend WithEvents varlabel5 As System.Windows.Forms.Label
    Friend WithEvents var6combo As System.Windows.Forms.ComboBox
    Friend WithEvents vno6 As System.Windows.Forms.Label
    Friend WithEvents vartext6 As System.Windows.Forms.TextBox
    Friend WithEvents varlabel6 As System.Windows.Forms.Label
    Friend WithEvents var7combo As System.Windows.Forms.ComboBox
    Friend WithEvents vno7 As System.Windows.Forms.Label
    Friend WithEvents vartext7 As System.Windows.Forms.TextBox
    Friend WithEvents varlabel7 As System.Windows.Forms.Label
    Friend WithEvents var8combo As System.Windows.Forms.ComboBox
    Friend WithEvents vno8 As System.Windows.Forms.Label
    Friend WithEvents vartext8 As System.Windows.Forms.TextBox
    Friend WithEvents varlabel8 As System.Windows.Forms.Label
    Friend WithEvents var9combo As System.Windows.Forms.ComboBox
    Friend WithEvents vno9 As System.Windows.Forms.Label
    Friend WithEvents vartext9 As System.Windows.Forms.TextBox
    Friend WithEvents varlabel9 As System.Windows.Forms.Label
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.varlabel1 = New System.Windows.Forms.Label
        Me.vartext1 = New System.Windows.Forms.TextBox
        Me.VScroll = New System.Windows.Forms.VScrollBar
        Me.vno1 = New System.Windows.Forms.Label
        Me.var1combo = New System.Windows.Forms.ComboBox
        Me.var2combo = New System.Windows.Forms.ComboBox
        Me.vno2 = New System.Windows.Forms.Label
        Me.vartext2 = New System.Windows.Forms.TextBox
        Me.varlabel2 = New System.Windows.Forms.Label
        Me.var3combo = New System.Windows.Forms.ComboBox
        Me.vno3 = New System.Windows.Forms.Label
        Me.vartext3 = New System.Windows.Forms.TextBox
        Me.varlabel3 = New System.Windows.Forms.Label
        Me.var4combo = New System.Windows.Forms.ComboBox
        Me.vno4 = New System.Windows.Forms.Label
        Me.vartext4 = New System.Windows.Forms.TextBox
        Me.varlabel4 = New System.Windows.Forms.Label
        Me.var5combo = New System.Windows.Forms.ComboBox
        Me.vno5 = New System.Windows.Forms.Label
        Me.vartext5 = New System.Windows.Forms.TextBox
        Me.varlabel5 = New System.Windows.Forms.Label
        Me.var6combo = New System.Windows.Forms.ComboBox
        Me.vno6 = New System.Windows.Forms.Label
        Me.vartext6 = New System.Windows.Forms.TextBox
        Me.varlabel6 = New System.Windows.Forms.Label
        Me.var7combo = New System.Windows.Forms.ComboBox
        Me.vno7 = New System.Windows.Forms.Label
        Me.vartext7 = New System.Windows.Forms.TextBox
        Me.varlabel7 = New System.Windows.Forms.Label
        Me.var8combo = New System.Windows.Forms.ComboBox
        Me.vno8 = New System.Windows.Forms.Label
        Me.vartext8 = New System.Windows.Forms.TextBox
        Me.varlabel8 = New System.Windows.Forms.Label
        Me.var9combo = New System.Windows.Forms.ComboBox
        Me.vno9 = New System.Windows.Forms.Label
        Me.vartext9 = New System.Windows.Forms.TextBox
        Me.varlabel9 = New System.Windows.Forms.Label
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.vartext10 = New System.Windows.Forms.TextBox
        Me.vartext11 = New System.Windows.Forms.TextBox
        Me.var10combo = New System.Windows.Forms.ComboBox
        Me.var11combo = New System.Windows.Forms.ComboBox
        Me.vno10 = New System.Windows.Forms.Label
        Me.vno11 = New System.Windows.Forms.Label
        Me.varlabel10 = New System.Windows.Forms.Label
        Me.varlabel11 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'varlabel1
        '
        Me.varlabel1.Location = New System.Drawing.Point(8, 8)
        Me.varlabel1.Name = "varlabel1"
        Me.varlabel1.Size = New System.Drawing.Size(88, 16)
        Me.varlabel1.Text = "L"
        '
        'vartext1
        '
        Me.vartext1.Location = New System.Drawing.Point(96, 8)
        Me.vartext1.Name = "vartext1"
        Me.vartext1.Size = New System.Drawing.Size(117, 21)
        Me.vartext1.TabIndex = 44
        '
        'VScroll
        '
        Me.VScroll.LargeChange = 2
        Me.VScroll.Location = New System.Drawing.Point(216, 8)
        Me.VScroll.Maximum = 4
        Me.VScroll.Name = "VScroll"
        Me.VScroll.Size = New System.Drawing.Size(16, 258)
        Me.VScroll.TabIndex = 43
        '
        'vno1
        '
        Me.vno1.Location = New System.Drawing.Point(232, 8)
        Me.vno1.Name = "vno1"
        Me.vno1.Size = New System.Drawing.Size(56, 16)
        Me.vno1.Text = "Label2"
        Me.vno1.Visible = False
        '
        'var1combo
        '
        Me.var1combo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
        Me.var1combo.Location = New System.Drawing.Point(96, 8)
        Me.var1combo.Name = "var1combo"
        Me.var1combo.Size = New System.Drawing.Size(117, 22)
        Me.var1combo.TabIndex = 40
        '
        'var2combo
        '
        Me.var2combo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
        Me.var2combo.Location = New System.Drawing.Point(96, 32)
        Me.var2combo.Name = "var2combo"
        Me.var2combo.Size = New System.Drawing.Size(117, 22)
        Me.var2combo.TabIndex = 35
        '
        'vno2
        '
        Me.vno2.Location = New System.Drawing.Point(232, 32)
        Me.vno2.Name = "vno2"
        Me.vno2.Size = New System.Drawing.Size(56, 16)
        Me.vno2.Text = "Label2"
        Me.vno2.Visible = False
        '
        'vartext2
        '
        Me.vartext2.Location = New System.Drawing.Point(96, 32)
        Me.vartext2.Name = "vartext2"
        Me.vartext2.Size = New System.Drawing.Size(117, 21)
        Me.vartext2.TabIndex = 38
        '
        'varlabel2
        '
        Me.varlabel2.Location = New System.Drawing.Point(8, 32)
        Me.varlabel2.Name = "varlabel2"
        Me.varlabel2.Size = New System.Drawing.Size(88, 16)
        Me.varlabel2.Text = "L"
        '
        'var3combo
        '
        Me.var3combo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
        Me.var3combo.Location = New System.Drawing.Point(96, 56)
        Me.var3combo.Name = "var3combo"
        Me.var3combo.Size = New System.Drawing.Size(117, 22)
        Me.var3combo.TabIndex = 30
        '
        'vno3
        '
        Me.vno3.Location = New System.Drawing.Point(232, 56)
        Me.vno3.Name = "vno3"
        Me.vno3.Size = New System.Drawing.Size(56, 16)
        Me.vno3.Text = "Label2"
        Me.vno3.Visible = False
        '
        'vartext3
        '
        Me.vartext3.Location = New System.Drawing.Point(96, 56)
        Me.vartext3.Name = "vartext3"
        Me.vartext3.Size = New System.Drawing.Size(117, 21)
        Me.vartext3.TabIndex = 33
        '
        'varlabel3
        '
        Me.varlabel3.Location = New System.Drawing.Point(8, 56)
        Me.varlabel3.Name = "varlabel3"
        Me.varlabel3.Size = New System.Drawing.Size(88, 16)
        Me.varlabel3.Text = "L"
        '
        'var4combo
        '
        Me.var4combo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
        Me.var4combo.Location = New System.Drawing.Point(96, 80)
        Me.var4combo.Name = "var4combo"
        Me.var4combo.Size = New System.Drawing.Size(117, 22)
        Me.var4combo.TabIndex = 25
        '
        'vno4
        '
        Me.vno4.Location = New System.Drawing.Point(232, 80)
        Me.vno4.Name = "vno4"
        Me.vno4.Size = New System.Drawing.Size(56, 16)
        Me.vno4.Text = "Label2"
        Me.vno4.Visible = False
        '
        'vartext4
        '
        Me.vartext4.Location = New System.Drawing.Point(96, 80)
        Me.vartext4.Name = "vartext4"
        Me.vartext4.Size = New System.Drawing.Size(117, 21)
        Me.vartext4.TabIndex = 28
        '
        'varlabel4
        '
        Me.varlabel4.Location = New System.Drawing.Point(8, 80)
        Me.varlabel4.Name = "varlabel4"
        Me.varlabel4.Size = New System.Drawing.Size(88, 16)
        Me.varlabel4.Text = "L"
        '
        'var5combo
        '
        Me.var5combo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
        Me.var5combo.Location = New System.Drawing.Point(96, 104)
        Me.var5combo.Name = "var5combo"
        Me.var5combo.Size = New System.Drawing.Size(117, 22)
        Me.var5combo.TabIndex = 20
        '
        'vno5
        '
        Me.vno5.Location = New System.Drawing.Point(232, 104)
        Me.vno5.Name = "vno5"
        Me.vno5.Size = New System.Drawing.Size(56, 16)
        Me.vno5.Text = "Label2"
        Me.vno5.Visible = False
        '
        'vartext5
        '
        Me.vartext5.Location = New System.Drawing.Point(96, 104)
        Me.vartext5.Name = "vartext5"
        Me.vartext5.Size = New System.Drawing.Size(117, 21)
        Me.vartext5.TabIndex = 23
        '
        'varlabel5
        '
        Me.varlabel5.Location = New System.Drawing.Point(8, 104)
        Me.varlabel5.Name = "varlabel5"
        Me.varlabel5.Size = New System.Drawing.Size(88, 16)
        Me.varlabel5.Text = "L"
        '
        'var6combo
        '
        Me.var6combo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
        Me.var6combo.Location = New System.Drawing.Point(96, 128)
        Me.var6combo.Name = "var6combo"
        Me.var6combo.Size = New System.Drawing.Size(117, 22)
        Me.var6combo.TabIndex = 15
        '
        'vno6
        '
        Me.vno6.Location = New System.Drawing.Point(232, 128)
        Me.vno6.Name = "vno6"
        Me.vno6.Size = New System.Drawing.Size(56, 16)
        Me.vno6.Text = "Label2"
        Me.vno6.Visible = False
        '
        'vartext6
        '
        Me.vartext6.Location = New System.Drawing.Point(96, 128)
        Me.vartext6.Name = "vartext6"
        Me.vartext6.Size = New System.Drawing.Size(117, 21)
        Me.vartext6.TabIndex = 18
        '
        'varlabel6
        '
        Me.varlabel6.Location = New System.Drawing.Point(8, 128)
        Me.varlabel6.Name = "varlabel6"
        Me.varlabel6.Size = New System.Drawing.Size(88, 16)
        Me.varlabel6.Text = "L"
        '
        'var7combo
        '
        Me.var7combo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
        Me.var7combo.Location = New System.Drawing.Point(96, 152)
        Me.var7combo.Name = "var7combo"
        Me.var7combo.Size = New System.Drawing.Size(117, 22)
        Me.var7combo.TabIndex = 10
        '
        'vno7
        '
        Me.vno7.Location = New System.Drawing.Point(232, 152)
        Me.vno7.Name = "vno7"
        Me.vno7.Size = New System.Drawing.Size(56, 16)
        Me.vno7.Text = "Label2"
        Me.vno7.Visible = False
        '
        'vartext7
        '
        Me.vartext7.Location = New System.Drawing.Point(96, 152)
        Me.vartext7.Name = "vartext7"
        Me.vartext7.Size = New System.Drawing.Size(117, 21)
        Me.vartext7.TabIndex = 13
        '
        'varlabel7
        '
        Me.varlabel7.Location = New System.Drawing.Point(8, 152)
        Me.varlabel7.Name = "varlabel7"
        Me.varlabel7.Size = New System.Drawing.Size(88, 16)
        Me.varlabel7.Text = "L"
        '
        'var8combo
        '
        Me.var8combo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
        Me.var8combo.Location = New System.Drawing.Point(96, 176)
        Me.var8combo.Name = "var8combo"
        Me.var8combo.Size = New System.Drawing.Size(117, 22)
        Me.var8combo.TabIndex = 5
        '
        'vno8
        '
        Me.vno8.Location = New System.Drawing.Point(232, 176)
        Me.vno8.Name = "vno8"
        Me.vno8.Size = New System.Drawing.Size(56, 16)
        Me.vno8.Text = "Label2"
        Me.vno8.Visible = False
        '
        'vartext8
        '
        Me.vartext8.Location = New System.Drawing.Point(96, 176)
        Me.vartext8.Name = "vartext8"
        Me.vartext8.Size = New System.Drawing.Size(117, 21)
        Me.vartext8.TabIndex = 8
        '
        'varlabel8
        '
        Me.varlabel8.Location = New System.Drawing.Point(8, 176)
        Me.varlabel8.Name = "varlabel8"
        Me.varlabel8.Size = New System.Drawing.Size(88, 16)
        Me.varlabel8.Text = "L"
        '
        'var9combo
        '
        Me.var9combo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
        Me.var9combo.Location = New System.Drawing.Point(96, 199)
        Me.var9combo.Name = "var9combo"
        Me.var9combo.Size = New System.Drawing.Size(117, 22)
        Me.var9combo.TabIndex = 0
        '
        'vno9
        '
        Me.vno9.Location = New System.Drawing.Point(232, 200)
        Me.vno9.Name = "vno9"
        Me.vno9.Size = New System.Drawing.Size(56, 16)
        Me.vno9.Text = "Label2"
        Me.vno9.Visible = False
        '
        'vartext9
        '
        Me.vartext9.Location = New System.Drawing.Point(96, 199)
        Me.vartext9.Name = "vartext9"
        Me.vartext9.Size = New System.Drawing.Size(117, 21)
        Me.vartext9.TabIndex = 3
        '
        'varlabel9
        '
        Me.varlabel9.Location = New System.Drawing.Point(8, 200)
        Me.varlabel9.Name = "varlabel9"
        Me.varlabel9.Size = New System.Drawing.Size(88, 16)
        Me.varlabel9.Text = "L"
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.Add(Me.MenuItem1)
        Me.MainMenu1.MenuItems.Add(Me.MenuItem3)
        '
        'MenuItem1
        '
        Me.MenuItem1.Text = "Save                    "
        '
        'MenuItem3
        '
        Me.MenuItem3.Text = "Cancel"
        '
        'vartext10
        '
        Me.vartext10.Location = New System.Drawing.Point(96, 222)
        Me.vartext10.Name = "vartext10"
        Me.vartext10.Size = New System.Drawing.Size(117, 21)
        Me.vartext10.TabIndex = 46
        '
        'vartext11
        '
        Me.vartext11.Location = New System.Drawing.Point(96, 245)
        Me.vartext11.Name = "vartext11"
        Me.vartext11.Size = New System.Drawing.Size(117, 21)
        Me.vartext11.TabIndex = 47
        '
        'var10combo
        '
        Me.var10combo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
        Me.var10combo.Location = New System.Drawing.Point(96, 222)
        Me.var10combo.Name = "var10combo"
        Me.var10combo.Size = New System.Drawing.Size(117, 22)
        Me.var10combo.TabIndex = 48
        '
        'var11combo
        '
        Me.var11combo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown
        Me.var11combo.Location = New System.Drawing.Point(96, 245)
        Me.var11combo.Name = "var11combo"
        Me.var11combo.Size = New System.Drawing.Size(117, 22)
        Me.var11combo.TabIndex = 49
        '
        'vno10
        '
        Me.vno10.Location = New System.Drawing.Point(232, 222)
        Me.vno10.Name = "vno10"
        Me.vno10.Size = New System.Drawing.Size(56, 16)
        Me.vno10.Text = "Label2"
        Me.vno10.Visible = False
        '
        'vno11
        '
        Me.vno11.Location = New System.Drawing.Point(233, 248)
        Me.vno11.Name = "vno11"
        Me.vno11.Size = New System.Drawing.Size(56, 16)
        Me.vno11.Text = "Label2"
        Me.vno11.Visible = False
        '
        'varlabel10
        '
        Me.varlabel10.Location = New System.Drawing.Point(8, 222)
        Me.varlabel10.Name = "varlabel10"
        Me.varlabel10.Size = New System.Drawing.Size(88, 16)
        Me.varlabel10.Text = "L"
        '
        'varlabel11
        '
        Me.varlabel11.Location = New System.Drawing.Point(8, 244)
        Me.varlabel11.Name = "varlabel11"
        Me.varlabel11.Size = New System.Drawing.Size(88, 16)
        Me.varlabel11.Text = "L"
        '
        'frmEditPoint
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.ClientSize = New System.Drawing.Size(240, 268)
        Me.ControlBox = False
        Me.Controls.Add(Me.varlabel11)
        Me.Controls.Add(Me.varlabel10)
        Me.Controls.Add(Me.vno11)
        Me.Controls.Add(Me.vno10)
        Me.Controls.Add(Me.vartext11)
        Me.Controls.Add(Me.vartext10)
        Me.Controls.Add(Me.vno9)
        Me.Controls.Add(Me.vartext9)
        Me.Controls.Add(Me.varlabel9)
        Me.Controls.Add(Me.vno8)
        Me.Controls.Add(Me.vartext8)
        Me.Controls.Add(Me.varlabel8)
        Me.Controls.Add(Me.vno7)
        Me.Controls.Add(Me.vartext7)
        Me.Controls.Add(Me.varlabel7)
        Me.Controls.Add(Me.vno6)
        Me.Controls.Add(Me.vartext6)
        Me.Controls.Add(Me.varlabel6)
        Me.Controls.Add(Me.vno5)
        Me.Controls.Add(Me.vartext5)
        Me.Controls.Add(Me.varlabel5)
        Me.Controls.Add(Me.vno4)
        Me.Controls.Add(Me.vartext4)
        Me.Controls.Add(Me.varlabel4)
        Me.Controls.Add(Me.vno3)
        Me.Controls.Add(Me.vartext3)
        Me.Controls.Add(Me.varlabel3)
        Me.Controls.Add(Me.vno2)
        Me.Controls.Add(Me.vartext2)
        Me.Controls.Add(Me.varlabel2)
        Me.Controls.Add(Me.vno1)
        Me.Controls.Add(Me.VScroll)
        Me.Controls.Add(Me.vartext1)
        Me.Controls.Add(Me.varlabel1)
        Me.Controls.Add(Me.var11combo)
        Me.Controls.Add(Me.var10combo)
        Me.Controls.Add(Me.var9combo)
        Me.Controls.Add(Me.var8combo)
        Me.Controls.Add(Me.var7combo)
        Me.Controls.Add(Me.var6combo)
        Me.Controls.Add(Me.var5combo)
        Me.Controls.Add(Me.var4combo)
        Me.Controls.Add(Me.var3combo)
        Me.Controls.Add(Me.var2combo)
        Me.Controls.Add(Me.var1combo)
        Me.KeyPreview = True
        Me.Menu = Me.MainMenu1
        Me.Name = "frmEditPoint"
        Me.Text = "Edit Point"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub frmEditPoint_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim a As Integer
        Dim flag As Boolean

        'Otherwise indicate that now we don't want anychanges
        inprocessedit = True

        VScroll.Minimum = 1
        If vars > 11 Then
            VScroll.Maximum = vars - 10
        Else
            VScroll.Maximum = 1
        End If
        VScroll.LargeChange = 1
        VScroll.Value = 1
        firstvar = 1

        If Not back_to_edit And Not back_to_main Then

            'First - clear the array of responses
            flag = False
            For a = 1 To vars
                responses(a) = ""
                If vcarry(a) = True Then flag = True
            Next a

            'Second - if some variables carry or if a continuation shot - then get the last record and fill them in
            If flag Or shottype = -8 Then
                Dim points As New SqlCeConnection("datasource =" + options(1))
                Dim sqlpoints As SqlCeCommand = points.CreateCommand
                points.Open()
                sqlpoints.CommandText = "SELECT * FROM " & options(22) & " ORDER BY Recno DESC"
                Dim rdr As SqlCeDataReader = sqlpoints.ExecuteReader()
                Do While rdr.Read
                    For a = 1 To vars
                        If vcarry(a) = True Or shottype = -8 Then
                            If Not IsDBNull(rdr.Item(varlist(a))) Then responses(a) = rdr.Item(varlist(a)).ToString
                        End If
                    Next a
                    Exit Do
                Loop
                rdr.Close()
                points.Close()
            End If

            'Two prime - if special case of barcode scan of unit-id
            ' then put that in the correct fields
            If (bar_code_pre_scan And bar_code_unit_id <> "" And shottype <> -8) Then
                If InStr(bar_code_unit_id, "-") <> 0 Then
                    Dim unit As String = leftstr(bar_code_unit_id, InStr(bar_code_unit_id, "-") - 1)
                    Dim id As String = Mid(bar_code_unit_id, InStr(bar_code_unit_id, "-") + 1)
                    For a = 1 To vars
                        If LCase(varlist(a)) = "unit" Then
                            responses(a) = unit
                            last_unit = unit
                        End If
                        If LCase(varlist(a)) = "id" Then
                            responses(a) = id
                        End If
                    Next
                End If
            End If

            'Third - handle limit checking if this is not a continuation shot
            If shottype = -1 And use_limitchecking Then
                Call get_unit_info()
            ElseIf shottype = -1 Then
                'Default to the last unit recorded
                Call get_last_unit_info()
            End If

            'Three prime - need to put back the scanned UNIT ID because
            ' depending on what unit checking just did they may be overridden
            If (bar_code_pre_scan And bar_code_unit_id <> "" And shottype <> -8) Then
                If InStr(bar_code_unit_id, "-") <> 0 Then
                    Dim unit As String = leftstr(bar_code_unit_id, InStr(bar_code_unit_id, "-") - 1)
                    Dim id As String = Mid(bar_code_unit_id, InStr(bar_code_unit_id, "-") + 1)
                    For a = 1 To vars
                        If LCase(varlist(a)) = "unit" Then
                            responses(a) = unit
                            last_unit = unit
                        End If
                        If LCase(varlist(a)) = "id" Then
                            responses(a) = id
                        End If
                    Next
                End If
            End If

            'Fourth - fill in the defaults
            For a = 1 To vars
                Select Case UCase(varlist(a))
                    Case "BARCODE"
                        responses(a) = ""
                    Case "X"
                        responses(a) = FormatNumber(currentxp, 3)
                    Case "Y"
                        responses(a) = FormatNumber(currentyp, 3)
                    Case "Z"
                        responses(a) = FormatNumber(currentzp, 3)
                    Case "PRISM"
                        responses(a) = currentprismheight.ToString
                        putback = currentprismheight
                    Case "DAY", "DATE"
                        'responses(a) = Now.ToShortDateString
                        responses(a) = CStr(Now)
                    Case "TIME"
                        responses(a) = Now.ToLongTimeString
                    Case "DATUMNAME"
                        responses(a) = currentstationname
                    Case "DATUMX"
                        responses(a) = currentstationx.ToString
                    Case "DATUMY"
                        responses(a) = currentstationy.ToString
                    Case "DATUMZ"
                        responses(a) = currentstationz.ToString
                    Case "HANGLE"
                        responses(a) = currenthangle
                    Case "SLOPED"
                        responses(a) = FormatNumber(currentsloped, 3)
                    Case "VANGLE"
                        responses(a) = currentvangle
                    Case "SUFFIX"
                        If shottype = -1 Then
                            responses(a) = "0"
                        Else
                            If responses(a) <> "" Then
                                responses(a) = CStr(CInt(responses(a)) + 1)
                            Else
                                responses(a) = "0"
                            End If
                        End If
                    Case Else
                End Select
            Next a

            'Sixth - handle the increment fields
            If shottype <> -8 Then
                For a = 1 To vars
                    If (vincr(a) = True And UCase(varlist(a)) <> "ID") Or (UCase(varlist(a)) = "ID" And shottype = -1 And bar_code_unit_id = "") Then
                        If responses(a) <> "" Then
                            'check and make sure it is a number
                            If IsNumeric(responses(a)) Then
                                responses(a) = Trim(CStr(CInt(responses(a)) + 1))
                            End If
                        Else
                            responses(a) = "1"
                        End If
                    End If
                Next a
            End If

            'Seventh - if it was a button shot - override everthing with those defaults
            If button_no <> 0 Then
                For a = 1 To vars
                    If LCase(varlist(a)) = "id" Then
                        If LCase(button_values(button_no, a)) = "alpha" Then
                            If shottype <> -8 Then responses(a) = hash(5)
                        End If
                    ElseIf LCase(varlist(a)) = "unit" Then
                        'This is a special case brought up by David Braun.
                        'The idea is that a speed button puts you in a unit.
                        'First - if there is a units table, look for that unit
                        If button_values(button_no, a) <> "" Then
                            Dim units As New SqlCeConnection("Datasource='" + options(1) + "'")
                            Dim unitssql As SqlCeCommand = units.CreateCommand
                            units.Open()
                            unitssql.CommandText = "SELECT * FROM edm_units WHERE unit='" + button_values(button_no, a) + "'"
                            Dim rdr As SqlCeDataReader = unitssql.ExecuteReader
                            flag = False
                            Dim c As Integer
                            Dim b As Integer
                            Do While rdr.Read
                                flag = True
                                For c = 0 To rdr.FieldCount - 1
                                    For b = 1 To vars
                                        If LCase(rdr.GetName(c)) = LCase(varlist(b)) Then
                                            If Not IsDBNull(rdr.Item(c)) Then
                                                responses(b) = rdr.Item(c).ToString
                                            End If
                                            Exit For
                                        End If
                                    Next b
                                Next c
                                Exit Do
                            Loop
                            rdr.Close()
                            units.Close()

                            'TODO - could conditionally do this loop on flag (as set above)
                            ' what this means is that increments are only done when a valid unit is
                            ' found for the speed button.  Normally there should be a valid unit, but
                            ' if there is not, the program winds up bumping unit twice.
                            For b = 1 To vars
                                If vincr(b) = True Or (UCase(varlist(b)) = "ID" And shottype = -1) Then
                                    If responses(b) <> "" Then
                                        If IsNumeric(responses(b)) Then
                                            responses(b) = Trim(CStr(CInt(responses(b)) + 1))
                                        End If
                                    Else
                                        responses(b) = "1"
                                    End If
                                End If
                            Next b
                        End If
                    ElseIf button_values(button_no, a) <> "" Then
                        responses(a) = button_values(button_no, a)
                    End If
                Next a
            End If

        Else

            For a = 1 To vars
                Select Case UCase(varlist(a))
                    Case "PRISM"
                        putback = CSng(responses(a))
                        Exit For
                    Case Else
                End Select
            Next a

        End If

        skip_var = -1

        Call fill_fields(1)

        'It is to the side if first field is 
        If LCase(varlist(1)) <> "unit" Then
            Select Case vtype(1)
                Case 4, 5
                    VScroll.Focus()
                Case Else
                    vartext1.Focus()
            End Select
        Else
            VScroll.Focus()
        End If

        inprocessedit = False

        Cursor.Current = Cursors.Default

    End Sub

    Private Sub frmEditPoint_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated

    End Sub

    Sub fill_fields(ByVal firstvar As Integer)

        Dim varno As Integer

        'First, fill the field names
        varno = firstvar
        If skip_var < 0 Or skip_var <> 1 Then Call set_controls(varno, vars, varlabel1, vartext1, var1combo, vno1)
        If skip_var < 0 Or skip_var <> 2 Then Call set_controls(varno + 1, vars, varlabel2, vartext2, var2combo, vno2)
        If skip_var < 0 Or skip_var <> 3 Then Call set_controls(varno + 2, vars, varlabel3, vartext3, var3combo, vno3)
        If skip_var < 0 Or skip_var <> 4 Then Call set_controls(varno + 3, vars, varlabel4, vartext4, var4combo, vno4)
        If skip_var < 0 Or skip_var <> 5 Then Call set_controls(varno + 4, vars, varlabel5, vartext5, var5combo, vno5)
        If skip_var < 0 Or skip_var <> 6 Then Call set_controls(varno + 5, vars, varlabel6, vartext6, var6combo, vno6)
        If skip_var < 0 Or skip_var <> 7 Then Call set_controls(varno + 6, vars, varlabel7, vartext7, var7combo, vno7)
        If skip_var < 0 Or skip_var <> 8 Then Call set_controls(varno + 7, vars, varlabel8, vartext8, var8combo, vno8)
        If skip_var < 0 Or skip_var <> 9 Then Call set_controls(varno + 8, vars, varlabel9, vartext9, var9combo, vno9)
        If skip_var < 0 Or skip_var <> 10 Then Call set_controls(varno + 9, vars, varlabel10, vartext10, var10combo, vno10)
        If skip_var < 0 Or skip_var <> 11 Then Call set_controls(varno + 10, vars, varlabel11, vartext11, var11combo, vno11)

    End Sub

    Private Sub after_combo(ByRef varcombo As ComboBox, ByVal vno As Label)


    End Sub

    Private Sub get_last_unit_info()

        Dim a As Integer
        Dim b As Integer
        Dim flag As Boolean

        flag = False
        Dim units As New SqlCeConnection("Datasource='" + options(1) + "'")
        Dim unitssql As SqlCeCommand = units.CreateCommand
        unitssql.CommandText = "SELECT * FROM edm_units WHERE unit='" + last_unit + "'"
        units.Open()
        Dim rdr As SqlCeDataReader = unitssql.ExecuteReader
        Do While rdr.Read
            For a = 0 To rdr.FieldCount - 1
                For b = 1 To vars
                    If LCase(rdr.GetName(a)) = LCase(varlist(b)) Then
                        If Not IsDBNull(rdr.Item(a)) Then
                            responses(b) = rdr.Item(a).ToString
                        Else
                            responses(b) = ""
                        End If
                        Exit For
                    End If
                Next b
            Next a
            flag = True
            Exit Do
        Loop
        'If you are not in a unit, then clear the fields associated with a unit,
        'otherwise load the saved values.
        If Not flag Then
            For a = 0 To rdr.FieldCount - 1
                For b = 1 To vars
                    If LCase(rdr.GetName(a)) = LCase(varlist(b)) Then
                        responses(b) = ""
                        Exit For
                    End If
                Next b
            Next a
        End If
        rdr.Close()
        units.Close()

    End Sub

    Private Sub set_controls(ByVal varno As Integer, ByVal vars As Integer, ByRef varlabel As Label, ByRef vartext As TextBox, ByRef varcombo As ComboBox, ByRef vno As Label)

        If varno > vars Then
            varlabel.Visible = False
            vartext.Visible = False
            varcombo.Visible = False
        Else
            varlabel.Font = New Font(varlabel.Font.Name, varlabel.Font.Size, FontStyle.Bold)
            varlabel.ForeColor = Color.FromArgb(0, 0, 0)
            vno.Text = CStr(varno)
            varlabel.Text = varprompt(varno)
            varlabel.Visible = True
            Select Case LCase(varlist(varno))
                'Case "prism"
                '    vartext.Visible = False
                '    varcombo.Visible = True
                '    varcombo.Items.Clear()
                '    varcombo.ValueMember = "x"
                '   varcombo.Text = responses(varno)
                Case "unit"
                    vartext.Visible = True
                    varcombo.Visible = False
                    'varcombo.Items.Clear()
                    'varcombo.ValueMember = "x"
                    vartext.Text = responses(varno)
                    If responses(varno) <> last_unit And Not back_to_main And Not back_to_edit And use_limitchecking = True Then
                        varlabel.ForeColor = Color.FromArgb(255, 0, 0)
                        varlabel.Font = New Font(varlabel.Font.Name, varlabel.Font.Size, FontStyle.Bold)
                    End If
                Case "id"
                    vartext.Visible = False
                    varcombo.Visible = True
                    varcombo.Items.Clear()
                    varcombo.Items.Add(responses(varno))
                    varcombo.Items.Add(hash(vlen(varno)))
                    varcombo.ValueMember = "n"
                    varcombo.Text = responses(varno)
                    'Case "x", "y", "z"
                    '    vartext.Visible = False
                    '    varcombo.Visible = True
                    '    varcombo.Items.Clear()
                    '    varcombo.ValueMember = "x"
                    '    varcombo.Items.Add(responses(varno))
                    '    varcombo.Items.Add("Adjust")
                    '   varcombo.Text = responses(varno)
                    '   varcombo.ValueMember = "n"
                Case Else
                    'If vtype(varno) = 4 Then
                    ' vartext.Visible = False
                    ' varcombo.Visible = True
                    ' varcombo.Items.Clear()
                    ' varcombo.ValueMember = "x"
                    ' varcombo.Text = responses(varno)
                    'Else
                    vartext.Visible = True
                    varcombo.Visible = False
                    vartext.Text = responses(varno)
                    vartext.MaxLength = vlen(varno)
                    'End If
            End Select
        End If

    End Sub
    Private Sub fill_list(ByRef varno As Integer, ByRef varcombo As ComboBox)

        Dim a As Integer
        Dim t1 As String
        Dim t2 As String

        If varcombo.ValueMember = "n" Then Exit Sub

        varcombo.BeginUpdate()
        varcombo.Items.Clear()

        Select Case LCase(varlist(varno))
            Case "prism"

                Dim prisms As New SqlCeConnection("datasource='" + options(1) + "'")
                Dim sql As SqlCeCommand = prisms.CreateCommand

                Try
                    prisms.Open()
                    sql.CommandText = "SELECT * FROM edm_poles ORDER BY Name"
                    Dim rdr As SqlCeDataReader = sql.ExecuteReader
                    a = 0
                    Do While rdr.Read
                        varcombo.Items.Add(rdr.Item("name"))
                        a = a + 1
                        temp_prisms(a) = CSng(rdr.Item("height"))
                    Loop
                    rdr.Close()

                Catch sqlex As System.Data.SqlServerCe.SqlCeException
                    MsgBox("Error reading prisms table in " & options(1) & vbNewLine & sqlex.Message, MsgBoxStyle.Exclamation, "Issue")
                    Exit Try

                Finally
                    prisms.Close()

                End Try

            Case LCase("unit")

                Dim units As New SqlCeConnection("datasource='" + options(1) + "'")
                Dim sql As SqlCeCommand = units.CreateCommand

                Try
                    units.Open()
                    sql.CommandText = "SELECT unit FROM edm_units"
                    Dim rdr As SqlCeDataReader = sql.ExecuteReader
                    Do While rdr.Read
                        varcombo.Items.Add(rdr.Item("unit"))
                    Loop
                    rdr.Close()

                Catch sqlex As System.Data.SqlServerCe.SqlCeException
                    MsgBox("Error reading units table in " & options(1) & vbNewLine & sqlex.Message, MsgBoxStyle.Exclamation, "Issue")
                    Exit Try

                Finally
                    units.Close()

                End Try

            Case Else
                t1 = vmenu(varno)
                Do Until t1 = ""
                    a = InStr(t1, ",")
                    If a <> 0 Then
                        t2 = leftstr(t1, a - 1)
                        If t2 <> "" Then
                            varcombo.Items.Add(t2)
                        End If
                        t1 = Mid(t1, a + 1)
                    ElseIf t1 <> "" Then
                        varcombo.Items.Add(t1)
                        Exit Do
                    End If
                Loop

        End Select

        varcombo.EndUpdate()
        varcombo.ValueMember = "n"
        varcombo.Text = responses(varno)

    End Sub
    Sub get_unit_info()

        Dim a As Integer
        Dim b As Integer
        Dim flag As Boolean
        Dim d As Double

        Dim units As New SqlCeConnection("Datasource='" + options(1) + "'")
        Dim unitssql As SqlCeCommand = units.CreateCommand
        If options(17) = "" Then
            unitssql.CommandText = "SELECT minx,miny,maxx,maxy,centerx,centery,radius FROM edm_units"
        Else
            unitssql.CommandText = "SELECT minx,miny,maxx,maxy,centerx,centery,radius," + options(17) + " FROM edm_units"
        End If
        units.Open()

        Dim rdr As SqlCeDataReader = unitssql.ExecuteReader

        flag = False
        Do While rdr.Read
            If Not IsDBNull(rdr.Item("radius")) And Not IsDBNull(rdr.Item("centerx")) And Not IsDBNull(rdr.Item("centery")) Then
                If CDbl(rdr.Item("radius")) <> 0 Then
                    d = System.Math.Sqrt((CDbl(rdr.Item("centerx")) - currentxp) ^ 2 + (CDbl(rdr.Item("centery")) - currentyp) ^ 2)
                    If d < CDbl(rdr.Item("radius")) Then
                        flag = True
                        Exit Do
                    End If
                End If
            End If
            If Not IsDBNull(rdr.Item("minx")) And Not IsDBNull(rdr.Item("miny")) And Not IsDBNull(rdr.Item("maxx")) And Not IsDBNull(rdr.Item("maxy")) Then
                If CDbl(rdr.Item("minx")) <= currentxp And CDbl(rdr.Item("maxx")) >= currentxp And CDbl(rdr.Item("miny")) <= currentyp And CDbl(rdr.Item("maxy")) >= currentyp Then
                    flag = True
                    Exit Do
                End If
            End If
        Loop

        'If you are not in a unit, then clear the fields associated with a unit,
        'otherwise load the saved values.
        If Not flag Then
            For a = 0 To rdr.FieldCount - 1
                For b = 1 To vars
                    If LCase(rdr.GetName(a)) = LCase(varlist(b)) Then
                        responses(b) = ""
                        Exit For
                    End If
                Next b
            Next a
        Else
            For a = 0 To rdr.FieldCount - 1
                For b = 1 To vars
                    If LCase(rdr.GetName(a)) = LCase(varlist(b)) Then
                        If Not IsDBNull(rdr.Item(a)) Then
                            responses(b) = rdr.Item(a).ToString
                        Else
                            responses(b) = ""
                        End If
                        Exit For
                    End If
                Next b
            Next a
        End If
        rdr.Close()
        units.Close()
        rdr.Dispose()
        units.Dispose()

    End Sub

    Private Sub VScroll_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VScroll.ValueChanged

        Dim varno As Integer

        If inprocessedit Then Exit Sub

        inprocessedit = True

        Cursor.Current = Cursors.WaitCursor

        firstvar = VScroll.Value
        varno = firstvar

        Call fill_fields(varno)

        inprocessedit = False

        Cursor.Current = Cursors.Default

    End Sub

    Private Sub var1combo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles var1combo.SelectedIndexChanged

        If inprocessedit Then Exit Sub
        skip_var = 1
        editvar = 1
        Call after_combo(var1combo, vno1)
        skip_var = -1

    End Sub

    Private Sub var2combo_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles var2combo.SelectedIndexChanged

        If inprocessedit Then Exit Sub
        skip_var = 2
        editvar = 2
        Call after_combo(var2combo, vno2)
        skip_var = -1

    End Sub

    Private Sub var3combo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles var3combo.SelectedIndexChanged

        If inprocessedit Then Exit Sub
        skip_var = 3
        editvar = 3
        Call after_combo(var3combo, vno3)
        skip_var = -1

    End Sub

    Private Sub var4combo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles var4combo.SelectedIndexChanged

        If inprocessedit Then Exit Sub
        skip_var = 4
        editvar = 4
        Call after_combo(var4combo, vno4)
        skip_var = -1

    End Sub

    Private Sub var5combo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles var5combo.SelectedIndexChanged

        If inprocessedit Then Exit Sub
        skip_var = 5
        editvar = 5
        Call after_combo(var5combo, vno5)
        skip_var = -1

    End Sub

    Private Sub var6combo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles var6combo.SelectedIndexChanged

        If inprocessedit Then Exit Sub
        skip_var = 6
        editvar = 6
        Call after_combo(var6combo, vno6)
        skip_var = -1

    End Sub

    Private Sub var7combo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles var7combo.SelectedIndexChanged

        If inprocessedit Then Exit Sub
        skip_var = 7
        editvar = 7
        Call after_combo(var7combo, vno7)
        skip_var = -1

    End Sub

    Private Sub var8combo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles var8combo.SelectedIndexChanged

        If inprocessedit Then Exit Sub
        skip_var = 8
        editvar = 8
        Call after_combo(var8combo, vno8)
        skip_var = -1

    End Sub

    Private Sub var9combo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles var9combo.SelectedIndexChanged

        If inprocessedit Then Exit Sub
        skip_var = 9
        editvar = 9
        Call after_combo(var9combo, vno9)
        skip_var = -1

    End Sub

    Private Sub vartext1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext1.GotFocus

        If inprocessedit Then Exit Sub
        inprocessedit = True
        Call handle_focus(vartext1, CInt(vno1.Text))
        inprocessedit = False

    End Sub

    Private Sub vartext1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext1.LostFocus

    End Sub

    Private Sub vartext2_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext2.GotFocus

        If inprocessedit Then Exit Sub
        inprocessedit = True
        Call handle_focus(vartext2, CInt(vno2.Text))
        inprocessedit = False

    End Sub

    Private Sub vartext2_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext2.LostFocus

    End Sub

    Private Sub vartext3_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext3.GotFocus

        If inprocessedit Then Exit Sub
        inprocessedit = True
        Call handle_focus(vartext3, CInt(vno3.Text))
        inprocessedit = False

    End Sub

    Private Sub vartext3_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext3.LostFocus

    End Sub

    Private Sub vartext4_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext4.GotFocus

        If inprocessedit Then Exit Sub
        inprocessedit = True
        Call handle_focus(vartext4, CInt(vno4.Text))
        inprocessedit = False

    End Sub

    Private Sub vartext4_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext4.LostFocus

    End Sub

    Private Sub vartext5_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext5.GotFocus

        If inprocessedit Then Exit Sub
        inprocessedit = True
        Call handle_focus(vartext5, CInt(vno5.Text))
        inprocessedit = False

    End Sub

    Private Sub vartext5_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext5.LostFocus

    End Sub

    Private Sub vartext6_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext6.GotFocus

        If inprocessedit Then Exit Sub
        inprocessedit = True
        Call handle_focus(vartext6, CInt(vno6.Text))
        inprocessedit = False

    End Sub

    Private Sub vartext6_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext6.LostFocus

    End Sub

    Private Sub vartext7_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext7.GotFocus

        If inprocessedit Then Exit Sub
        inprocessedit = True
        Call handle_focus(vartext7, CInt(vno7.Text))
        inprocessedit = False

    End Sub

    Private Sub vartext7_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext7.LostFocus

    End Sub

    Private Sub vartext8_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext8.GotFocus

        If inprocessedit Then Exit Sub
        inprocessedit = True
        Call handle_focus(vartext8, CInt(vno8.Text))
        inprocessedit = False

    End Sub

    Private Sub vartext8_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext8.LostFocus

    End Sub

    Private Sub vartext9_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext9.GotFocus

        If inprocessedit Then Exit Sub
        inprocessedit = True
        Call handle_focus(vartext9, CInt(vno9.Text))
        inprocessedit = False

    End Sub

    Private Sub vartext9_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext9.LostFocus

    End Sub

    Private Sub var1combo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var1combo.GotFocus

        If inprocessedit Then Exit Sub
        Dim varno As Integer
        varno = CInt(vno1.Text)
        Call fill_list(varno, var1combo)

    End Sub

    Private Sub var2combo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var2combo.GotFocus

        If inprocessedit Then Exit Sub
        Dim varno As Integer
        varno = CInt(vno2.Text)
        Call fill_list(varno, var2combo)

    End Sub

    Private Sub var3combo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var3combo.GotFocus

        If inprocessedit Then Exit Sub
        Dim varno As Integer
        varno = CInt(vno3.Text)
        Call fill_list(varno, var3combo)

    End Sub

    Private Sub var4combo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var4combo.GotFocus

        If inprocessedit Then Exit Sub
        Dim varno As Integer
        varno = CInt(vno4.Text)
        Call fill_list(varno, var4combo)

    End Sub

    Private Sub var5combo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var5combo.GotFocus

        If inprocessedit Then Exit Sub
        Dim varno As Integer
        varno = CInt(vno5.Text)
        Call fill_list(varno, var5combo)

    End Sub

    Private Sub var6combo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var6combo.GotFocus

        If inprocessedit Then Exit Sub
        Dim varno As Integer
        varno = CInt(vno6.Text)
        Call fill_list(varno, var6combo)

    End Sub

    Private Sub var7combo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var7combo.GotFocus

        If inprocessedit Then Exit Sub
        Dim varno As Integer
        varno = CInt(vno7.Text)
        Call fill_list(varno, var7combo)

    End Sub

    Private Sub var8combo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var8combo.GotFocus

        If inprocessedit Then Exit Sub
        Dim varno As Integer
        varno = CInt(vno8.Text)
        Call fill_list(varno, var8combo)

    End Sub

    Private Sub var9combo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var9combo.GotFocus

        If inprocessedit Then Exit Sub
        Dim varno As Integer
        varno = CInt(vno9.Text)
        Call fill_list(varno, var9combo)

    End Sub

    Private Sub var1combo_DataSourceChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles var1combo.DataSourceChanged

    End Sub

    Private Sub var1combo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var1combo.LostFocus

        If LCase(varlist(CInt(vno1.Text))) = "prism" Then
            If var1combo.SelectedIndex <> -1 Then
                responses(CInt(vno1.Text)) = CStr(temp_prisms(var1combo.SelectedIndex))
                var1combo.Text = CStr(temp_prisms(var1combo.SelectedIndex))
            Else
                responses(CInt(vno1.Text)) = var1combo.Text
            End If
        Else
            responses(CInt(vno1.Text)) = var1combo.Text
        End If

    End Sub

    Private Sub var2combo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var2combo.LostFocus

        If LCase(varlist(CInt(vno2.Text))) = "prism" Then
            If var2combo.SelectedIndex <> -1 Then
                responses(CInt(vno2.Text)) = CStr(temp_prisms(var2combo.SelectedIndex))
                var2combo.Text = CStr(temp_prisms(var2combo.SelectedIndex))
            Else
                responses(CInt(vno2.Text)) = var2combo.Text
            End If
        Else
            responses(CInt(vno2.Text)) = var2combo.Text
        End If

    End Sub

    Private Sub var3combo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var3combo.LostFocus

        If LCase(varlist(CInt(vno3.Text))) = "prism" Then
            If var3combo.SelectedIndex <> -1 Then
                responses(CInt(vno3.Text)) = CStr(temp_prisms(var3combo.SelectedIndex))
                var3combo.Text = CStr(temp_prisms(var3combo.SelectedIndex))
            Else
                responses(CInt(vno3.Text)) = var3combo.Text
            End If
        Else
            responses(CInt(vno3.Text)) = var3combo.Text
        End If

    End Sub

    Private Sub var4combo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var4combo.LostFocus

        If LCase(varlist(CInt(vno4.Text))) = "prism" Then
            If var1combo.SelectedIndex <> -1 Then
                responses(CInt(vno4.Text)) = CStr(temp_prisms(var4combo.SelectedIndex))
                var4combo.Text = CStr(temp_prisms(var4combo.SelectedIndex))
            Else
                responses(CInt(vno4.Text)) = var4combo.Text
            End If
        Else
            responses(CInt(vno4.Text)) = var4combo.Text
        End If

    End Sub

    Private Sub var5combo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var5combo.LostFocus

        If LCase(varlist(CInt(vno5.Text))) = "prism" Then
            If var1combo.SelectedIndex <> -1 Then
                responses(CInt(vno5.Text)) = CStr(temp_prisms(var5combo.SelectedIndex))
                var5combo.Text = CStr(temp_prisms(var5combo.SelectedIndex))
            Else
                responses(CInt(vno5.Text)) = var5combo.Text
            End If
        Else
            responses(CInt(vno5.Text)) = var5combo.Text
        End If

    End Sub

    Private Sub var6combo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var6combo.LostFocus

        If LCase(varlist(CInt(vno6.Text))) = "prism" Then
            If var6combo.SelectedIndex <> -1 Then
                responses(CInt(vno6.Text)) = CStr(temp_prisms(var6combo.SelectedIndex))
                var6combo.Text = CStr(temp_prisms(var6combo.SelectedIndex))
            Else
                responses(CInt(vno6.Text)) = var6combo.Text
            End If
        Else
            responses(CInt(vno6.Text)) = var6combo.Text
        End If

    End Sub

    Private Sub var7combo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var7combo.LostFocus

        If LCase(varlist(CInt(vno7.Text))) = "prism" Then
            If var7combo.SelectedIndex <> -1 Then
                responses(CInt(vno7.Text)) = CStr(temp_prisms(var7combo.SelectedIndex))
                var7combo.Text = CStr(temp_prisms(var7combo.SelectedIndex))
            Else
                responses(CInt(vno7.Text)) = var7combo.Text
            End If
        Else
            responses(CInt(vno7.Text)) = var7combo.Text
        End If

    End Sub

    Private Sub var8combo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var8combo.LostFocus

        If LCase(varlist(CInt(vno8.Text))) = "prism" Then
            If var8combo.SelectedIndex <> -1 Then
                responses(CInt(vno8.Text)) = CStr(temp_prisms(var8combo.SelectedIndex))
                var8combo.Text = CStr(temp_prisms(var8combo.SelectedIndex))
            Else
                responses(CInt(vno8.Text)) = var8combo.Text
            End If
        Else
            responses(CInt(vno8.Text)) = var8combo.Text
        End If

    End Sub

    Private Sub var9combo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var9combo.LostFocus

        If LCase(varlist(CInt(vno9.Text))) = "prism" Then
            If var9combo.SelectedIndex <> -1 Then
                responses(CInt(vno9.Text)) = CStr(temp_prisms(var9combo.SelectedIndex))
                var9combo.Text = CStr(temp_prisms(var9combo.SelectedIndex))
            Else
                responses(CInt(vno9.Text)) = var9combo.Text
            End If
        Else
            responses(CInt(vno9.Text)) = var9combo.Text
        End If

    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click

        Dim a As Integer
        Dim b As Integer
        Dim flag As Boolean
        Dim recno As Integer
        Dim inidata(1, 2) As String
        Dim iniclass As String
        Dim wstatus As Integer

        Application.DoEvents()

        inprocessedit = True        'This seems to fix the bug wherein lostfocus events don't happen first
        Me.Focus()
        inprocessedit = False

        Cursor.Current = Cursors.WaitCursor
        Dim points As New SqlCeConnection("Datasource='" + options(1) + "'")
        Dim pointssql As SqlCeCommand = points.CreateCommand
        pointssql.CommandText = "SELECT "
        For a = 1 To vars
            pointssql.CommandText = pointssql.CommandText & varlist(a) & ","
        Next a
        pointssql.CommandText = pointssql.CommandText & "Recno FROM " & options(22) & " ORDER BY Recno DESC"
        points.Open()
        Dim rdr As SqlCeDataReader = pointssql.ExecuteReader

        'Do a quick check for invalid data
        For a = 1 To vars
            If responses(a) = "" And varrequired(a) = True Then
                Cursor.Current = Cursors.Default
                MsgBox("A value is required for field " & varlist(a), MsgBoxStyle.Information, "EDM-Mobile")
                rdr.Close()
                points.Close()
                Exit Sub
            End If
            If responses(a) <> "" Then
                Select Case rdr.GetDataTypeName(a - 1)
                    Case "DateTime"
                        If Not IsDate(responses(a)) Then
                            Cursor.Current = Cursors.Default
                            MsgBox("Invalid data or time entered for " + varlist(a), MsgBoxStyle.Information, "EDM-Mobile")
                            rdr.Close()
                            points.Close()
                            Exit Sub
                        End If
                    Case "NVarChar"
                        If Len(responses(a)) > vlen(a) Then
                            Cursor.Current = Cursors.Default
                            MsgBox("Value entered for " + varlist(a) + " is too long to fit in the field.", MsgBoxStyle.Information, "EDM-Mobile")
                            rdr.Close()
                            points.Close()
                            Exit Sub
                        End If
                    Case "Float"
                        If Not IsNumeric(responses(a)) Then
                            Cursor.Current = Cursors.Default
                            MsgBox("Value entered for " + varlist(a) + " must be a valid number.", MsgBoxStyle.Information, "EDM-Mobile")
                            rdr.Close()
                            points.Close()
                            Exit Sub
                        Else
                            Replace(responses(a), ",", "")
                        End If

                End Select
            End If
        Next a

        'Get the appropriate record number
        If back_to_main Then
            'if editing the last record, delete it and then re-insert it
            rdr.Read()
            recno = CInt(rdr.Item("Recno"))
            pointssql.CommandText = "DELETE FROM " + options(22) + " WHERE recno=" + CStr(recno)
            pointssql.ExecuteNonQuery()
        ElseIf rdr.Read Then
            'if adding a new record, then get the last recno and add 1
            If Not IsDBNull(rdr.Item("recno")) Then
                recno = CInt(rdr.Item("recno")) + 1
            Else
                recno = rdr.RecordsAffected + 1
            End If
        Else
            recno = 1
        End If

        'if this is the first record, make sure recno = 1
        If recno = 0 Then recno = 1

        If Not back_to_edit Then

            pointssql.CommandText = "INSERT INTO " + options(22) + " (recno"
            For a = 1 To vars
                pointssql.CommandText = pointssql.CommandText + "," + varlist(a)
            Next
            pointssql.CommandText = pointssql.CommandText + ") VALUES (" + CStr(recno)

            For a = 1 To vars
                If responses(a) = "" Then
                    Select Case rdr.GetDataTypeName(rdr.GetOrdinal(varlist(a)))
                        Case "NVarChar"
                            pointssql.CommandText = pointssql.CommandText + ",' '"
                        Case Else
                            pointssql.CommandText = pointssql.CommandText + ",0"
                    End Select
                Else
                    Select Case rdr.GetDataTypeName(rdr.GetOrdinal(varlist(a)))
                        Case "DateTime"
                            pointssql.CommandText = pointssql.CommandText + ",'" + responses(a) + "'"
                            'If UCase(varlist(a)) = "TIME" Then
                            ' pointssql.CommandText = pointssql.CommandText + ",'" + FormatDateTime(CDate(responses(a)), vbShortTime) + "'"
                            'Else
                            'pointssql.CommandText = pointssql.CommandText + ",'" + responses(a) + "'"
                            'End If
                        Case "NVarChar"
                            pointssql.CommandText = pointssql.CommandText & ",'" & responses(a) + "'"
                        Case Else
                            pointssql.CommandText = pointssql.CommandText & "," & Replace(responses(a), ",", "")
                    End Select
                End If
                If LCase(varlist(a)) = "id" Then last_id = responses(a)
                If LCase(varlist(a)) = "unit" Then last_unit = responses(a)
                If LCase(varlist(a)) = "suffix" Then last_suffix = responses(a)
            Next a
            pointssql.CommandText = pointssql.CommandText + ")"

            pointssql.ExecuteNonQuery()

            'frmDatagrid.pointgrid.Rows = frmDatagrid.pointgrid.Rows + 1
            'frmDatagrid.pointgrid.TextMatrix(frmDatagrid.pointgrid.Rows - 1, 0) = recno
            'For a = 1 To vars
            'If Not IsNull(rspoints(varlist(a))) Then frmDatagrid.pointgrid.TextMatrix(frmDatagrid.pointgrid.Rows - 1, a) = rspoints(varlist(a))
            'Next a

            returnedstatus = "Record no. " & CStr(recno) & " - Object no. " & last_unit & "-" & last_id & "(" & last_suffix & ")"

        End If

        'update values in limit checking
        If shottype = -1 And back_to_edit = False And last_suffix = "0" Then
            Dim unitssql As SqlCeCommand = points.CreateCommand
            unitssql.CommandText = "SELECT unit FROM edm_units WHERE unit='" + last_unit + "'"
            Dim r As Object
            r = unitssql.ExecuteScalar()

            If LCase(CStr(r)) = LCase(last_unit) Then
                unitssql.CommandText = "SELECT * FROM edm_units WHERE unit='" + last_unit + "'"
                Dim unitrdr As SqlCeDataReader = unitssql.ExecuteReader
                unitssql.CommandText = "UPDATE edm_units SET "
                flag = False
                For a = 0 To unitrdr.FieldCount - 1
                    Select Case LCase(unitrdr.GetName(a))
                        Case "unit", "minx", "miny", "maxx", "maxy", "centerx", "centery", "radius"
                        Case Else
                            For b = 1 To vars
                                If LCase(varlist(b)) = LCase(unitrdr.GetName(a)) Then
                                    If LCase(varlist(b)) = "id" And Not IsNumeric(responses(b)) Then
                                        Exit For
                                    End If
                                    If flag = False Then
                                        unitssql.CommandText = unitssql.CommandText + unitrdr.GetName(a) + "='" + responses(b) + "'"
                                        flag = True
                                    Else
                                        unitssql.CommandText = unitssql.CommandText + "," + unitrdr.GetName(a) + "='" + responses(b) + "'"
                                    End If
                                    Exit For
                                End If
                            Next
                    End Select
                Next
                If flag Then
                    unitssql.CommandText = unitssql.CommandText + " WHERE unit='" + last_unit + "'"
                    unitssql.ExecuteNonQuery()
                End If
                unitrdr.Close()
            Else
                'otherwise run an insert query
                Dim unitrdr As SqlCeDataReader = unitssql.ExecuteReader
                unitssql.CommandText = "INSERT INTO edm_units (unit,id,minx,miny,maxx,maxy,centerx,centery,radius"
                flag = False
                For a = 0 To unitrdr.FieldCount - 1
                    Select Case LCase(unitrdr.GetName(a))
                        Case "unit", "id", "minx", "miny", "maxx", "maxy", "centerx", "centery", "radius"
                        Case Else
                            unitssql.CommandText = unitssql.CommandText + "," + unitrdr.GetName(a)
                    End Select
                Next
                unitssql.CommandText = unitssql.CommandText + ") VALUES ('" + last_unit + "','" + last_id + "',0,0,0,0,0,0,0"
                If flag Then
                    For a = 0 To unitrdr.FieldCount - 1
                        Select Case LCase(unitrdr.GetName(a))
                            Case "unit", "id", "minx", "miny", "maxx", "maxy", "centerx", "centery", "radius"
                            Case Else
                                For b = 1 To vars
                                    If LCase(varlist(b)) = LCase(unitrdr.GetName(a)) Then
                                        unitssql.CommandText = unitssql.CommandText + "," + responses(b)
                                        Exit For
                                    End If
                                Next
                        End Select
                    Next
                End If
                unitssql.CommandText = unitssql.CommandText + ")"
                unitssql.ExecuteNonQuery()
                unitrdr.Close()
            End If
            unitssql.Dispose()
        End If

        'See if any menus need to be updated
        For a = 1 To vars
            If vtype(a) = 4 And responses(a) <> "" Then
                If InStr(LCase(vmenu(a)), "," + LCase(responses(a)) + ",") = 0 Then
                    If LCase(Microsoft.VisualBasic.Left(vmenu(a), Len(responses(a)) + 1)) <> LCase(responses(a)) + "," Then
                        If LCase(vmenu(a)) <> LCase(responses(a)) Then
                            If LCase(Microsoft.VisualBasic.Right(vmenu(a), Len(responses(a)) + 1)) <> "," + LCase(responses(a)) Then
                                If vmenu(a) = "" Then
                                    vmenu(a) = responses(a)
                                Else
                                    vmenu(a) = vmenu(a) & "," & responses(a)
                                End If
                                iniclass = "[" + varlist(a) + "]"
                                inidata(1, 1) = "Menu"
                                inidata(1, 2) = vmenu(a)
                                Call WriteIni(cfgfile, iniclass, inidata, wstatus)
                            End If
                        End If
                    End If
                End If
            End If
        Next a

        ' If a pre-shot bar code scan was done and not a continuation shot,
        ' bump the ID by one to prepare for the next shot
        If bar_code_unit_id <> "" And shottype <> -8 Then
            If InStr(bar_code_unit_id, "-") <> 0 Then
                a = InStr(bar_code_unit_id, "-")
                Dim unit As String = leftstr(bar_code_unit_id, a - 1)
                Dim id As String = Mid(bar_code_unit_id, a + 1)
                If IsNumeric(id) Then id = Str(Val(id) + 1)
                bar_code_unit_id = unit + "-" + id
            End If
        End If

        pointssql.Dispose()

        rdr.Close()
        points.Close()

        rdr.Dispose()
        points.Dispose()

        Cursor.Current = Cursors.Default

        Me.DialogResult = Windows.Forms.DialogResult.OK

    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click

        Me.DialogResult = Windows.Forms.DialogResult.Cancel

    End Sub

    Private Sub var10combo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var10combo.GotFocus

        If inprocessedit Then Exit Sub
        Dim varno As Integer
        varno = CInt(vno10.Text)
        Call fill_list(varno, var10combo)

    End Sub

    Private Sub var10combo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var10combo.LostFocus

        If LCase(varlist(CInt(vno10.Text))) = "prism" Then
            If var10combo.SelectedIndex <> -1 Then
                responses(CInt(vno10.Text)) = CStr(temp_prisms(var10combo.SelectedIndex))
                var10combo.Text = CStr(temp_prisms(var10combo.SelectedIndex))
            Else
                responses(CInt(vno10.Text)) = var10combo.Text
            End If
        Else
            responses(CInt(vno10.Text)) = var10combo.Text
        End If


    End Sub

    Private Sub var11combo_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var11combo.GotFocus

        If inprocessedit Then Exit Sub
        Dim varno As Integer
        varno = CInt(vno11.Text)
        Call fill_list(varno, var11combo)

    End Sub

    Private Sub var11combo_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles var11combo.LostFocus

        If LCase(varlist(CInt(vno11.Text))) = "prism" Then
            If var11combo.SelectedIndex <> -1 Then
                responses(CInt(vno11.Text)) = CStr(temp_prisms(var11combo.SelectedIndex))
                var11combo.Text = CStr(temp_prisms(var11combo.SelectedIndex))
            Else
                responses(CInt(vno11.Text)) = var11combo.Text
            End If
        Else
            responses(CInt(vno11.Text)) = var11combo.Text
        End If

    End Sub

    Private Sub var10combo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles var10combo.SelectedIndexChanged

        If inprocessedit Then Exit Sub
        skip_var = 10
        editvar = 10
        Call after_combo(var10combo, vno10)
        skip_var = -1

    End Sub

    Private Sub var11combo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles var11combo.SelectedIndexChanged

        If inprocessedit Then Exit Sub
        skip_var = 11
        editvar = 11
        Call after_combo(var11combo, vno11)
        skip_var = -1

    End Sub

    Private Sub vartext10_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext10.GotFocus

        If inprocessedit Then Exit Sub
        inprocessedit = True
        Call handle_focus(vartext10, CInt(vno10.Text))
        inprocessedit = False

    End Sub

    Private Sub vartext10_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext10.LostFocus

    End Sub

    Private Sub vartext11_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext11.GotFocus

        If inprocessedit Then Exit Sub
        inprocessedit = True
        Call handle_focus(vartext11, CInt(vno11.Text))
        inprocessedit = False

    End Sub

    Private Sub vartext11_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext11.LostFocus
    End Sub

    Private Sub frmEditPoint_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If (e.KeyCode = System.Windows.Forms.Keys.Up) Then
            'Up
        End If
        If (e.KeyCode = System.Windows.Forms.Keys.Down) Then
            'Down
        End If
        If (e.KeyCode = System.Windows.Forms.Keys.Left) Then
            'Left
        End If
        If (e.KeyCode = System.Windows.Forms.Keys.Right) Then
            'Right
        End If
        If (e.KeyCode = System.Windows.Forms.Keys.Enter) Then
            'Enter
        End If

    End Sub
    Private Sub fill_menu(ByRef formtofill As frmMenulist, ByVal variableno As Integer)

        Dim a As Integer
        Dim t1 As String
        Dim t2 As String

        formtofill.ListBox1.BeginUpdate()
        formtofill.ListBox1.Items.Clear()

        Select Case LCase(varlist(variableno))
            Case "prism"

                Dim prisms As New SqlCeConnection("datasource='" + options(1) + "'")
                Dim sql As SqlCeCommand = prisms.CreateCommand

                Try
                    prisms.Open()
                    sql.CommandText = "SELECT * FROM edm_poles ORDER BY Name"
                    Dim rdr As SqlCeDataReader = sql.ExecuteReader
                    a = 0
                    Do While rdr.Read
                        formtofill.ListBox1.Items.Add(rdr.Item("name"))
                        a = a + 1
                        temp_prisms(a) = CSng(rdr.Item("height"))
                    Loop
                    rdr.Close()

                Catch sqlex As System.Data.SqlServerCe.SqlCeException
                    MsgBox("Error reading prisms table in " & options(1) & vbNewLine & sqlex.Message, MsgBoxStyle.Exclamation, "Issue")
                    Exit Try

                Finally
                    prisms.Close()

                End Try

            Case LCase("unit")

                Dim units As New SqlCeConnection("datasource='" + options(1) + "'")
                Dim sql As SqlCeCommand = units.CreateCommand

                Try
                    units.Open()
                    sql.CommandText = "SELECT unit FROM edm_units"
                    Dim rdr As SqlCeDataReader = sql.ExecuteReader
                    Do While rdr.Read
                        formtofill.ListBox1.Items.Add(rdr.Item("unit"))
                    Loop
                    rdr.Close()

                Catch sqlex As System.Data.SqlServerCe.SqlCeException
                    MsgBox("Error reading units table in " & options(1) & vbNewLine & sqlex.Message, MsgBoxStyle.Exclamation, "Issue")
                    Exit Try

                Finally
                    units.Close()

                End Try

            Case Else
                t1 = vmenu(variableno)
                Do Until t1 = ""
                    a = InStr(t1, ",")
                    If a <> 0 Then
                        t2 = leftstr(t1, a - 1)
                        If t2 <> "" Then
                            formtofill.ListBox1.Items.Add(t2)
                        End If
                        t1 = Mid(t1, a + 1)
                    ElseIf t1 <> "" Then
                        formtofill.ListBox1.Items.Add(t1)
                        Exit Do
                    End If
                Loop

        End Select

        formtofill.ListBox1.EndUpdate()
        formtofill.ListBox1.SelectedItem = responses(variableno)

    End Sub
    Private Sub handle_focus(ByRef vartext As TextBox, ByVal varno As Integer)

        If vtype(varno) = 4 Or LCase(varlist(varno)) = "prism" Or LCase(varlist(varno)) = "unit" Or LCase(varlist(varno)) = "x" Or LCase(varlist(varno)) = "y" Or LCase(varlist(varno)) = "z" Then

            Dim adj As Single
            Dim a As Integer
            Dim b As Integer
            Dim u As String
            Dim flag As Boolean

            Select Case LCase(varlist(varno))
                Case "prism"
                    Dim menulist As New frmMenulist
                    Call fill_menu(menulist, varno)
                    If menulist.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
                        If MsgBox("Update Z as well?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "EDM-Mobile") = MsgBoxResult.Yes Then
                            For a = 1 To npoles
                                If polename(a) = menuresult Then
                                    adj = CSng(poledata(a, 1))
                                    vartext.Text = adj.ToString
                                    responses(varno) = vartext.Text
                                    Exit For
                                End If
                            Next a
                            For a = 1 To vars
                                If LCase(varlist(a)) = "z" Then
                                    responses(a) = Format(CDbl(responses(a)) + putback - adj, "########0.000")
                                    If CInt(vno1.Text) = a Then
                                        vartext1.Text = responses(a)
                                    ElseIf CInt(vno2.Text) = a Then
                                        vartext2.Text = responses(a)
                                    ElseIf CInt(vno3.Text) = a Then
                                        vartext3.Text = responses(a)
                                    ElseIf CInt(vno4.Text) = a Then
                                        vartext4.Text = responses(a)
                                    ElseIf CInt(vno5.Text) = a Then
                                        vartext5.Text = responses(a)
                                    ElseIf CInt(vno6.Text) = a Then
                                        vartext6.Text = responses(a)
                                    ElseIf CInt(vno7.Text) = a Then
                                        vartext7.Text = responses(a)
                                    ElseIf CInt(vno8.Text) = a Then
                                        vartext8.Text = responses(a)
                                    ElseIf CInt(vno9.Text) = a Then
                                        vartext9.Text = responses(a)
                                    ElseIf CInt(vno10.Text) = a Then
                                        vartext10.Text = responses(a)
                                    ElseIf CInt(vno11.Text) = a Then
                                        vartext11.Text = responses(a)
                                    End If
                                    Exit For
                                End If
                            Next a
                            putback = adj
                        End If
                    End If

                Case "unit"
                    Dim menulist As New frmMenulist
                    Call fill_menu(menulist, varno)
                    If menulist.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
                        responses(varno) = menuresult
                        vartext.Text = menuresult
                        u = LCase(menuresult)

                        'First - if there is a units table, look for that unit
                        Dim units As New SqlCeConnection("Datasource='" + options(1) + "'")
                        Dim unitssql As SqlCeCommand = units.CreateCommand
                        units.Open()
                        unitssql.CommandText = "SELECT * FROM edm_units WHERE unit='" + u + "'"
                        Dim rdr As SqlCeDataReader = unitssql.ExecuteReader
                        flag = False
                        Do While rdr.Read
                            flag = True
                            For a = 0 To rdr.FieldCount - 1
                                For b = 1 To vars
                                    If LCase(rdr.GetName(a)) = LCase(varlist(b)) Then
                                        If Not IsDBNull(rdr.Item(a)) Then
                                            responses(b) = rdr.Item(a).ToString
                                        End If
                                        Exit For
                                    End If
                                Next b
                            Next a
                            Exit Do
                        Loop
                        If Not flag Then
                            'If no units table - then just clear out unit fields
                            For a = 0 To rdr.FieldCount - 1
                                For b = 1 To vars
                                    If LCase(rdr.GetName(a)) = LCase(varlist(b)) Then
                                        If LCase(varlist(b)) <> "unit" Then responses(b) = ""
                                        Exit For
                                    End If
                                Next b
                            Next a
                        End If
                        rdr.Close()
                        units.Close()

                        For a = 1 To vars
                            If vincr(a) = True Or (UCase(varlist(a)) = "ID" And shottype = -1) Then
                                If responses(a) <> "" Then
                                    If IsNumeric(responses(a)) Then
                                        responses(a) = Trim(CStr(CInt(responses(a)) + 1))
                                    End If
                                Else
                                    responses(a) = "1"
                                End If
                            End If
                        Next a

                        inprocessedit = False
                        Dim e As System.EventArgs = System.EventArgs.Empty
                        Call VScroll_ValueChanged(Me, e)
                    End If

                Case "x"
                    Dim xyzadjform As New frmXYZadj
                    actualvar = varno
                    xyzadjform.direction1.Text = "West"
                    xyzadjform.direction2.Text = "East"
                    xyzadjform.Text = "X Adjust"
                    xyzadjform.direction1.Left = xyzadjform.text1.Left
                    xyzadjform.direction1.Top = xyzadjform.text1.Top - xyzadjform.direction1.Height
                    xyzadjform.direction2.Left = xyzadjform.units.Left
                    xyzadjform.direction2.Top = xyzadjform.direction1.Top
                    inprocessedit = True
                    If xyzadjform.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        If responses(actualvar) = "" Then
                            responses(actualvar) = measurementadjustment.ToString
                        Else
                            responses(actualvar) = CStr(CDbl(responses(actualvar)) + measurementadjustment)
                        End If
                    End If
                    vartext.Text = responses(actualvar)
                    inprocessedit = False

                Case "y"
                    Dim xyzadjform As New frmXYZadj
                    actualvar = varno
                    xyzadjform.direction1.Text = "South"
                    xyzadjform.direction2.Text = "North"
                    xyzadjform.Text = "Y Adjust"
                    xyzadjform.direction1.Top = xyzadjform.text1.Top - xyzadjform.direction1.Height
                    xyzadjform.direction2.Top = xyzadjform.direction1.Top - xyzadjform.direction2.Height
                    xyzadjform.direction1.Left = xyzadjform.Width \ 2 - xyzadjform.direction1.Width \ 2
                    xyzadjform.direction2.Left = xyzadjform.Width \ 2 - xyzadjform.direction2.Width \ 2
                    inprocessedit = True
                    If xyzadjform.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        If responses(actualvar) = "" Then
                            responses(actualvar) = measurementadjustment.ToString
                        Else
                            responses(actualvar) = CStr(CDbl(responses(actualvar)) + measurementadjustment)
                        End If
                    End If
                    vartext.Text = responses(actualvar)
                    inprocessedit = False

                Case "z"
                    Dim xyzadjform As New frmXYZadj
                    actualvar = varno
                    xyzadjform.direction1.Text = "Down"
                    xyzadjform.direction2.Text = "Up"
                    xyzadjform.Text = "Z Adjust"
                    xyzadjform.direction1.Top = xyzadjform.text1.Top - xyzadjform.direction1.Height
                    xyzadjform.direction2.Top = xyzadjform.direction1.Top - xyzadjform.direction2.Height
                    xyzadjform.direction1.Left = xyzadjform.Width \ 2 - xyzadjform.direction1.Width \ 2
                    xyzadjform.direction2.Left = xyzadjform.Width \ 2 - xyzadjform.direction1.Width \ 2
                    inprocessedit = True
                    If xyzadjform.ShowDialog() = Windows.Forms.DialogResult.OK Then
                        If responses(actualvar) = "" Then
                            responses(actualvar) = measurementadjustment.ToString
                        Else
                            responses(actualvar) = CStr(CDbl(responses(actualvar)) + measurementadjustment)
                        End If
                    End If
                    vartext.Text = responses(actualvar)
                    inprocessedit = False

                Case Else
                    Dim menulist As New frmMenulist
                    Call fill_menu(menulist, varno)
                    If menulist.ShowDialog() <> Windows.Forms.DialogResult.Cancel Then
                        responses(varno) = menuresult
                        vartext.Text = menuresult
                    End If

            End Select
            VScroll.Focus()

        End If

    End Sub

    Private Sub vartext1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles vartext1.TextChanged

        If inprocessedit Then Exit Sub
        If vtype(CInt(vno1.Text)) = 2 And vartext1.Text <> "" Then
            If Not IsNumeric(vartext1.Text) Then
                MsgBox("Value must be numeric for this field.", MsgBoxStyle.Information, "EDM-Mobile")
                Exit Sub
            End If
        End If
        responses(CInt(vno1.Text)) = vartext1.Text
        If InStr(vartext1.Text, "-") <> 0 Then
            inprocessedit = True
            delay(0.1)
            Application.DoEvents()
            Dim a As Integer = InStr(vartext1.Text, "-")
            Dim unit As String = leftstr(vartext1.Text, a - 1)
            Dim id As String = Mid(vartext1.Text, a + 1)
            For a = 1 To vars
                If LCase(varlist(a)) = "unit" Then responses(a) = unit
                If LCase(varlist(a)) = "id" Then responses(a) = id
            Next
            firstvar = VScroll.Value
            Dim varno As Integer = firstvar
            responses(CInt(vno1.Text)) = vartext1.Text
            fill_fields(varno)
            vartext1.SelectAll()
            inprocessedit = False
        End If

    End Sub

    Private Sub vartext8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles vartext8.TextChanged

        If inprocessedit Then Exit Sub
        If vtype(CInt(vno8.Text)) = 2 And vartext8.Text <> "" Then
            If Not IsNumeric(vartext8.Text) Then
                MsgBox("Value must be numeric for this field.", MsgBoxStyle.Information, "EDM-Mobile")
                Exit Sub
            End If
        End If
        responses(CInt(vno8.Text)) = vartext8.Text

    End Sub

    Private Sub vartext9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles vartext9.TextChanged

        If inprocessedit Then Exit Sub
        If vtype(CInt(vno9.Text)) = 2 And vartext9.Text <> "" Then
            If Not IsNumeric(vartext9.Text) Then
                MsgBox("Value must be numeric for this field.", MsgBoxStyle.Information, "EDM-Mobile")
                Exit Sub
            End If
        End If
        responses(CInt(vno9.Text)) = vartext9.Text

    End Sub

    Private Sub vartext2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext2.TextChanged

        If inprocessedit Then Exit Sub
        If vtype(CInt(vno2.Text)) = 2 And vartext2.Text <> "" Then
            If Not IsNumeric(vartext2.Text) Then
                MsgBox("Value must be numeric for this field.", MsgBoxStyle.Information, "EDM-Mobile")
                Exit Sub
            End If
        End If
        responses(CInt(vno2.Text)) = vartext2.Text

    End Sub

    Private Sub vartext3_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext3.TextChanged

        If inprocessedit Then Exit Sub
        If vtype(CInt(vno3.Text)) = 2 And vartext3.Text <> "" Then
            If Not IsNumeric(vartext3.Text) Then
                MsgBox("Value must be numeric for this field.", MsgBoxStyle.Information, "EDM-Mobile")
                Exit Sub
            End If
        End If
        responses(CInt(vno3.Text)) = vartext3.Text

    End Sub

    Private Sub vartext4_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext4.TextChanged

        If inprocessedit Then Exit Sub
        If vtype(CInt(vno4.Text)) = 2 And vartext4.Text <> "" Then
            If Not IsNumeric(vartext4.Text) Then
                MsgBox("Value must be numeric for this field.", MsgBoxStyle.Information, "EDM-Mobile")
                Exit Sub
            End If
        End If
        responses(CInt(vno4.Text)) = vartext4.Text

    End Sub

    Private Sub vartext5_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext5.TextChanged

        If inprocessedit Then Exit Sub
        If vtype(CInt(vno5.Text)) = 2 And vartext5.Text <> "" Then
            If Not IsNumeric(vartext5.Text) Then
                MsgBox("Value must be numeric for this field.", MsgBoxStyle.Information, "EDM-Mobile")
                Exit Sub
            End If
        End If
        responses(CInt(vno5.Text)) = vartext5.Text

    End Sub

    Private Sub vartext6_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext6.TextChanged

        If inprocessedit Then Exit Sub
        If vtype(CInt(vno6.Text)) = 2 And vartext6.Text <> "" Then
            If Not IsNumeric(vartext6.Text) Then
                MsgBox("Value must be numeric for this field.", MsgBoxStyle.Information, "EDM-Mobile")
                Exit Sub
            End If
        End If
        responses(CInt(vno6.Text)) = vartext6.Text

    End Sub

    Private Sub vartext7_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext7.TextChanged

        If inprocessedit Then Exit Sub
        If vtype(CInt(vno7.Text)) = 2 And vartext7.Text <> "" Then
            If Not IsNumeric(vartext7.Text) Then
                MsgBox("Value must be numeric for this field.", MsgBoxStyle.Information, "EDM-Mobile")
                Exit Sub
            End If
        End If
        responses(CInt(vno7.Text)) = vartext7.Text

    End Sub

    Private Sub vartext10_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext10.TextChanged

        If inprocessedit Then Exit Sub
        If vtype(CInt(vno10.Text)) = 2 And vartext10.Text <> "" Then
            If Not IsNumeric(vartext10.Text) Then
                MsgBox("Value must be numeric for this field.", MsgBoxStyle.Information, "EDM-Mobile")
                Exit Sub
            End If
        End If
        responses(CInt(vno10.Text)) = vartext10.Text

    End Sub

    Private Sub vartext11_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles vartext11.TextChanged

        If inprocessedit Then Exit Sub
        If vtype(CInt(vno11.Text)) = 2 And vartext11.Text <> "" Then
            If Not IsNumeric(vartext11.Text) Then
                MsgBox("Value must be numeric for this field.", MsgBoxStyle.Information, "EDM-Mobile")
                Exit Sub
            End If
        End If
        responses(CInt(vno11.Text)) = vartext11.Text

    End Sub
End Class
