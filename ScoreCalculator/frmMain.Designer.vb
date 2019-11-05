<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.menuStrip = New System.Windows.Forms.MenuStrip()
        Me.PreProcessToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.InProcessToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PostProcessToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PostProcessToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.ScoreManagerToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.WindowsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CascadeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CloseAllToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'menuStrip
        '
        Me.menuStrip.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.menuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.PreProcessToolStripMenuItem, Me.InProcessToolStripMenuItem, Me.PostProcessToolStripMenuItem, Me.WindowsToolStripMenuItem})
        Me.menuStrip.Location = New System.Drawing.Point(0, 0)
        Me.menuStrip.Name = "menuStrip"
        Me.menuStrip.Size = New System.Drawing.Size(932, 28)
        Me.menuStrip.TabIndex = 1
        Me.menuStrip.Text = "MenuStrip1"
        '
        'PreProcessToolStripMenuItem
        '
        Me.PreProcessToolStripMenuItem.Name = "PreProcessToolStripMenuItem"
        Me.PreProcessToolStripMenuItem.Size = New System.Drawing.Size(95, 24)
        Me.PreProcessToolStripMenuItem.Text = "P&re Process"
        '
        'InProcessToolStripMenuItem
        '
        Me.InProcessToolStripMenuItem.Name = "InProcessToolStripMenuItem"
        Me.InProcessToolStripMenuItem.Size = New System.Drawing.Size(86, 24)
        Me.InProcessToolStripMenuItem.Text = "&In Process"
        '
        'PostProcessToolStripMenuItem
        '
        Me.PostProcessToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.PostProcessToolStripMenuItem1, Me.ScoreManagerToolStripMenuItem})
        Me.PostProcessToolStripMenuItem.Name = "PostProcessToolStripMenuItem"
        Me.PostProcessToolStripMenuItem.Size = New System.Drawing.Size(101, 24)
        Me.PostProcessToolStripMenuItem.Text = "P&ost Process"
        '
        'PostProcessToolStripMenuItem1
        '
        Me.PostProcessToolStripMenuItem1.Name = "PostProcessToolStripMenuItem1"
        Me.PostProcessToolStripMenuItem1.Size = New System.Drawing.Size(216, 26)
        Me.PostProcessToolStripMenuItem1.Text = "&Post Process"
        '
        'ScoreManagerToolStripMenuItem
        '
        Me.ScoreManagerToolStripMenuItem.Name = "ScoreManagerToolStripMenuItem"
        Me.ScoreManagerToolStripMenuItem.Size = New System.Drawing.Size(216, 26)
        Me.ScoreManagerToolStripMenuItem.Text = "&Score Manager"
        '
        'WindowsToolStripMenuItem
        '
        Me.WindowsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CascadeToolStripMenuItem, Me.CloseAllToolStripMenuItem})
        Me.WindowsToolStripMenuItem.Name = "WindowsToolStripMenuItem"
        Me.WindowsToolStripMenuItem.Size = New System.Drawing.Size(82, 24)
        Me.WindowsToolStripMenuItem.Text = "&Windows"
        '
        'CascadeToolStripMenuItem
        '
        Me.CascadeToolStripMenuItem.Name = "CascadeToolStripMenuItem"
        Me.CascadeToolStripMenuItem.Size = New System.Drawing.Size(142, 26)
        Me.CascadeToolStripMenuItem.Text = "&Cascade"
        '
        'CloseAllToolStripMenuItem
        '
        Me.CloseAllToolStripMenuItem.Name = "CloseAllToolStripMenuItem"
        Me.CloseAllToolStripMenuItem.Size = New System.Drawing.Size(142, 26)
        Me.CloseAllToolStripMenuItem.Text = "C&lose All"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(932, 503)
        Me.Controls.Add(Me.menuStrip)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.menuStrip
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Score Calculator"
        Me.menuStrip.ResumeLayout(False)
        Me.menuStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents menuStrip As MenuStrip
    Friend WithEvents PreProcessToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents InProcessToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents PostProcessToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents WindowsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CascadeToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CloseAllToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents PostProcessToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents ScoreManagerToolStripMenuItem As ToolStripMenuItem
End Class
