<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Public Class frmPlots
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPlots))
        Me.mainMenu1 = New System.Windows.Forms.MainMenu
        Me.no_action = New System.Windows.Forms.MenuItem
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuExit = New System.Windows.Forms.MenuItem
        Me.plot = New System.Windows.Forms.PictureBox
        Me.zoom_in = New System.Windows.Forms.PictureBox
        Me.zoom_out = New System.Windows.Forms.PictureBox
        Me.full_extent = New System.Windows.Forms.PictureBox
        Me.grid = New System.Windows.Forms.PictureBox
        Me.info = New System.Windows.Forms.PictureBox
        Me.infobox = New System.Windows.Forms.TextBox
        Me.XY = New System.Windows.Forms.PictureBox
        Me.XZ = New System.Windows.Forms.PictureBox
        Me.YZ = New System.Windows.Forms.PictureBox
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog
        Me.SuspendLayout()
        '
        'mainMenu1
        '
        Me.mainMenu1.MenuItems.Add(Me.no_action)
        '
        'no_action
        '
        Me.no_action.MenuItems.Add(Me.MenuItem1)
        Me.no_action.MenuItems.Add(Me.MenuItem2)
        Me.no_action.MenuItems.Add(Me.MenuExit)
        Me.no_action.Text = "Menu"
        '
        'MenuItem1
        '
        Me.MenuItem1.Text = "Select Overlay File"
        '
        'MenuItem2
        '
        Me.MenuItem2.Text = "Show/Hide Overlay"
        '
        'MenuExit
        '
        Me.MenuExit.Text = "Close"
        '
        'plot
        '
        Me.plot.BackColor = System.Drawing.Color.Black
        Me.plot.Location = New System.Drawing.Point(10, 6)
        Me.plot.Name = "plot"
        Me.plot.Size = New System.Drawing.Size(220, 220)
        '
        'zoom_in
        '
        Me.zoom_in.BackColor = System.Drawing.Color.White
        Me.zoom_in.Image = CType(resources.GetObject("zoom_in.Image"), System.Drawing.Image)
        Me.zoom_in.Location = New System.Drawing.Point(10, 231)
        Me.zoom_in.Name = "zoom_in"
        Me.zoom_in.Size = New System.Drawing.Size(34, 34)
        Me.zoom_in.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        '
        'zoom_out
        '
        Me.zoom_out.Image = CType(resources.GetObject("zoom_out.Image"), System.Drawing.Image)
        Me.zoom_out.Location = New System.Drawing.Point(47, 231)
        Me.zoom_out.Name = "zoom_out"
        Me.zoom_out.Size = New System.Drawing.Size(34, 34)
        Me.zoom_out.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        '
        'full_extent
        '
        Me.full_extent.Image = CType(resources.GetObject("full_extent.Image"), System.Drawing.Image)
        Me.full_extent.Location = New System.Drawing.Point(84, 231)
        Me.full_extent.Name = "full_extent"
        Me.full_extent.Size = New System.Drawing.Size(34, 34)
        Me.full_extent.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        '
        'grid
        '
        Me.grid.Image = CType(resources.GetObject("grid.Image"), System.Drawing.Image)
        Me.grid.Location = New System.Drawing.Point(121, 231)
        Me.grid.Name = "grid"
        Me.grid.Size = New System.Drawing.Size(34, 34)
        Me.grid.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        '
        'info
        '
        Me.info.Image = CType(resources.GetObject("info.Image"), System.Drawing.Image)
        Me.info.Location = New System.Drawing.Point(158, 231)
        Me.info.Name = "info"
        Me.info.Size = New System.Drawing.Size(34, 34)
        Me.info.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        '
        'infobox
        '
        Me.infobox.Location = New System.Drawing.Point(56, 48)
        Me.infobox.Multiline = True
        Me.infobox.Name = "infobox"
        Me.infobox.Size = New System.Drawing.Size(117, 74)
        Me.infobox.TabIndex = 8
        Me.infobox.Visible = False
        Me.infobox.WordWrap = False
        '
        'XY
        '
        Me.XY.Image = CType(resources.GetObject("XY.Image"), System.Drawing.Image)
        Me.XY.Location = New System.Drawing.Point(196, 231)
        Me.XY.Name = "XY"
        Me.XY.Size = New System.Drawing.Size(34, 34)
        Me.XY.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        '
        'XZ
        '
        Me.XZ.Image = CType(resources.GetObject("XZ.Image"), System.Drawing.Image)
        Me.XZ.Location = New System.Drawing.Point(110, 117)
        Me.XZ.Name = "XZ"
        Me.XZ.Size = New System.Drawing.Size(34, 34)
        Me.XZ.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.XZ.Visible = False
        '
        'YZ
        '
        Me.YZ.Image = CType(resources.GetObject("YZ.Image"), System.Drawing.Image)
        Me.YZ.Location = New System.Drawing.Point(118, 125)
        Me.YZ.Name = "YZ"
        Me.YZ.Size = New System.Drawing.Size(34, 34)
        Me.YZ.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.YZ.Visible = False
        '
        'OpenFileDialog
        '
        Me.OpenFileDialog.FileName = "OpenFileDialog"
        '
        'frmPlots
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(240, 268)
        Me.ControlBox = False
        Me.Controls.Add(Me.YZ)
        Me.Controls.Add(Me.XZ)
        Me.Controls.Add(Me.XY)
        Me.Controls.Add(Me.infobox)
        Me.Controls.Add(Me.info)
        Me.Controls.Add(Me.grid)
        Me.Controls.Add(Me.full_extent)
        Me.Controls.Add(Me.zoom_out)
        Me.Controls.Add(Me.zoom_in)
        Me.Controls.Add(Me.plot)
        Me.Menu = Me.mainMenu1
        Me.MinimizeBox = False
        Me.Name = "frmPlots"
        Me.Text = "Plotting"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents plot As System.Windows.Forms.PictureBox
    Friend WithEvents zoom_in As System.Windows.Forms.PictureBox
    Friend WithEvents no_action As System.Windows.Forms.MenuItem
    Friend WithEvents zoom_out As System.Windows.Forms.PictureBox
    Friend WithEvents full_extent As System.Windows.Forms.PictureBox
    Friend WithEvents grid As System.Windows.Forms.PictureBox
    Friend WithEvents info As System.Windows.Forms.PictureBox
    Friend WithEvents infobox As System.Windows.Forms.TextBox
    Friend WithEvents XY As System.Windows.Forms.PictureBox
    Friend WithEvents XZ As System.Windows.Forms.PictureBox
    Friend WithEvents YZ As System.Windows.Forms.PictureBox
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuExit As System.Windows.Forms.MenuItem
    Friend WithEvents OpenFileDialog As System.Windows.Forms.OpenFileDialog
End Class
