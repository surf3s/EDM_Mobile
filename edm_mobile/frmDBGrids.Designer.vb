<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class frmDBGrids
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Private mainMenu1 As System.Windows.Forms.MainMenu

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.mainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.dbgrid = New System.Windows.Forms.TabControl
        Me.pointpage = New System.Windows.Forms.TabPage
        Me.pointgrid = New System.Windows.Forms.DataGrid
        Me.datumpage = New System.Windows.Forms.TabPage
        Me.datumgrid = New System.Windows.Forms.DataGrid
        Me.unitspage = New System.Windows.Forms.TabPage
        Me.unitsgrid = New System.Windows.Forms.DataGrid
        Me.prismspage = New System.Windows.Forms.TabPage
        Me.prismsgrid = New System.Windows.Forms.DataGrid
        Me.bspoints = New System.Windows.Forms.BindingSource(Me.components)
        Me.bsdatums = New System.Windows.Forms.BindingSource(Me.components)
        Me.bsunits = New System.Windows.Forms.BindingSource(Me.components)
        Me.bsprisms = New System.Windows.Forms.BindingSource(Me.components)
        Me.dbgrid.SuspendLayout()
        Me.pointpage.SuspendLayout()
        Me.datumpage.SuspendLayout()
        Me.unitspage.SuspendLayout()
        Me.prismspage.SuspendLayout()
        CType(Me.bspoints, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.bsdatums, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.bsunits, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.bsprisms, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'mainMenu1
        '
        Me.mainMenu1.MenuItems.Add(Me.MenuItem1)
        Me.mainMenu1.MenuItems.Add(Me.MenuItem2)
        '
        'MenuItem1
        '
        Me.MenuItem1.Text = "Data"
        '
        'MenuItem2
        '
        Me.MenuItem2.Text = "Close"
        '
        'dbgrid
        '
        Me.dbgrid.Controls.Add(Me.pointpage)
        Me.dbgrid.Controls.Add(Me.datumpage)
        Me.dbgrid.Controls.Add(Me.unitspage)
        Me.dbgrid.Controls.Add(Me.prismspage)
        Me.dbgrid.Location = New System.Drawing.Point(0, 0)
        Me.dbgrid.Name = "dbgrid"
        Me.dbgrid.SelectedIndex = 0
        Me.dbgrid.Size = New System.Drawing.Size(240, 265)
        Me.dbgrid.TabIndex = 0
        '
        'pointpage
        '
        Me.pointpage.Controls.Add(Me.pointgrid)
        Me.pointpage.Location = New System.Drawing.Point(0, 0)
        Me.pointpage.Name = "pointpage"
        Me.pointpage.Size = New System.Drawing.Size(240, 242)
        Me.pointpage.Text = "Points"
        '
        'pointgrid
        '
        Me.pointgrid.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.pointgrid.Location = New System.Drawing.Point(0, 4)
        Me.pointgrid.Name = "pointgrid"
        Me.pointgrid.Size = New System.Drawing.Size(237, 235)
        Me.pointgrid.TabIndex = 0
        '
        'datumpage
        '
        Me.datumpage.Controls.Add(Me.datumgrid)
        Me.datumpage.Location = New System.Drawing.Point(0, 0)
        Me.datumpage.Name = "datumpage"
        Me.datumpage.Size = New System.Drawing.Size(232, 239)
        Me.datumpage.Text = "Datums"
        '
        'datumgrid
        '
        Me.datumgrid.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.datumgrid.Location = New System.Drawing.Point(2, 4)
        Me.datumgrid.Name = "datumgrid"
        Me.datumgrid.Size = New System.Drawing.Size(237, 235)
        Me.datumgrid.TabIndex = 1
        '
        'unitspage
        '
        Me.unitspage.Controls.Add(Me.unitsgrid)
        Me.unitspage.Location = New System.Drawing.Point(0, 0)
        Me.unitspage.Name = "unitspage"
        Me.unitspage.Size = New System.Drawing.Size(232, 239)
        Me.unitspage.Text = "Units"
        '
        'unitsgrid
        '
        Me.unitsgrid.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.unitsgrid.Location = New System.Drawing.Point(2, 4)
        Me.unitsgrid.Name = "unitsgrid"
        Me.unitsgrid.Size = New System.Drawing.Size(237, 235)
        Me.unitsgrid.TabIndex = 1
        '
        'prismspage
        '
        Me.prismspage.Controls.Add(Me.prismsgrid)
        Me.prismspage.Location = New System.Drawing.Point(0, 0)
        Me.prismspage.Name = "prismspage"
        Me.prismspage.Size = New System.Drawing.Size(232, 239)
        Me.prismspage.Text = "Prisms"
        '
        'prismsgrid
        '
        Me.prismsgrid.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.prismsgrid.Location = New System.Drawing.Point(2, 4)
        Me.prismsgrid.Name = "prismsgrid"
        Me.prismsgrid.Size = New System.Drawing.Size(237, 235)
        Me.prismsgrid.TabIndex = 1
        '
        'frmDBGrids
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(240, 268)
        Me.Controls.Add(Me.dbgrid)
        Me.Menu = Me.mainMenu1
        Me.Name = "frmDBGrids"
        Me.Text = "frmDBGrids"
        Me.dbgrid.ResumeLayout(False)
        Me.pointpage.ResumeLayout(False)
        Me.datumpage.ResumeLayout(False)
        Me.unitspage.ResumeLayout(False)
        Me.prismspage.ResumeLayout(False)
        CType(Me.bspoints, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.bsdatums, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.bsunits, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.bsprisms, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents dbgrid As System.Windows.Forms.TabControl
    Friend WithEvents pointpage As System.Windows.Forms.TabPage
    Friend WithEvents datumpage As System.Windows.Forms.TabPage
    Friend WithEvents unitspage As System.Windows.Forms.TabPage
    Friend WithEvents prismspage As System.Windows.Forms.TabPage
    Friend WithEvents pointgrid As System.Windows.Forms.DataGrid
    Friend WithEvents datumgrid As System.Windows.Forms.DataGrid
    Friend WithEvents unitsgrid As System.Windows.Forms.DataGrid
    Friend WithEvents prismsgrid As System.Windows.Forms.DataGrid
    Friend WithEvents bspoints As System.Windows.Forms.BindingSource
    Friend WithEvents bsdatums As System.Windows.Forms.BindingSource
    Friend WithEvents bsunits As System.Windows.Forms.BindingSource
    Friend WithEvents bsprisms As System.Windows.Forms.BindingSource
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
End Class
