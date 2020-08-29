<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmAllProcess
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAllProcess))
        Me.btnStart = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.txtFolderpath = New System.Windows.Forms.TextBox()
        Me.lblFolderPath = New System.Windows.Forms.Label()
        Me.grpFolderBrowse = New System.Windows.Forms.GroupBox()
        Me.chkbAutoProcess = New System.Windows.Forms.CheckBox()
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.folderBrowse = New System.Windows.Forms.FolderBrowserDialog()
        Me.lstError = New System.Windows.Forms.ListBox()
        Me.lblMainProgress = New System.Windows.Forms.Label()
        Me.grpFolderBrowse.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnStart
        '
        Me.btnStart.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStart.Location = New System.Drawing.Point(767, 89)
        Me.btnStart.Margin = New System.Windows.Forms.Padding(4)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(104, 34)
        Me.btnStart.TabIndex = 11
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.BackColor = System.Drawing.Color.SkyBlue
        Me.Panel2.Location = New System.Drawing.Point(14, 1)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(991, 12)
        Me.Panel2.TabIndex = 10
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BackColor = System.Drawing.Color.SkyBlue
        Me.Panel1.Location = New System.Drawing.Point(1, 1)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(13, 465)
        Me.Panel1.TabIndex = 9
        '
        'btnBrowse
        '
        Me.btnBrowse.Location = New System.Drawing.Point(852, 23)
        Me.btnBrowse.Margin = New System.Windows.Forms.Padding(4)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(100, 28)
        Me.btnBrowse.TabIndex = 2
        Me.btnBrowse.Text = "Browse"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'txtFolderpath
        '
        Me.txtFolderpath.Location = New System.Drawing.Point(97, 25)
        Me.txtFolderpath.Margin = New System.Windows.Forms.Padding(4)
        Me.txtFolderpath.Name = "txtFolderpath"
        Me.txtFolderpath.Size = New System.Drawing.Size(747, 22)
        Me.txtFolderpath.TabIndex = 1
        '
        'lblFolderPath
        '
        Me.lblFolderPath.AutoSize = True
        Me.lblFolderPath.Location = New System.Drawing.Point(8, 29)
        Me.lblFolderPath.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblFolderPath.Name = "lblFolderPath"
        Me.lblFolderPath.Size = New System.Drawing.Size(81, 17)
        Me.lblFolderPath.TabIndex = 0
        Me.lblFolderPath.Text = "Folder Path"
        '
        'grpFolderBrowse
        '
        Me.grpFolderBrowse.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpFolderBrowse.Controls.Add(Me.btnBrowse)
        Me.grpFolderBrowse.Controls.Add(Me.txtFolderpath)
        Me.grpFolderBrowse.Controls.Add(Me.lblFolderPath)
        Me.grpFolderBrowse.Location = New System.Drawing.Point(21, 17)
        Me.grpFolderBrowse.Margin = New System.Windows.Forms.Padding(4)
        Me.grpFolderBrowse.Name = "grpFolderBrowse"
        Me.grpFolderBrowse.Padding = New System.Windows.Forms.Padding(4)
        Me.grpFolderBrowse.Size = New System.Drawing.Size(963, 66)
        Me.grpFolderBrowse.TabIndex = 8
        Me.grpFolderBrowse.TabStop = False
        Me.grpFolderBrowse.Text = "Settings"
        '
        'chkbAutoProcess
        '
        Me.chkbAutoProcess.AutoSize = True
        Me.chkbAutoProcess.Location = New System.Drawing.Point(594, 96)
        Me.chkbAutoProcess.Name = "chkbAutoProcess"
        Me.chkbAutoProcess.Size = New System.Drawing.Size(166, 21)
        Me.chkbAutoProcess.TabIndex = 11
        Me.chkbAutoProcess.Text = "Auto Process All Step"
        Me.chkbAutoProcess.UseVisualStyleBackColor = True
        '
        'lblProgress
        '
        Me.lblProgress.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblProgress.Location = New System.Drawing.Point(21, 157)
        Me.lblProgress.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(963, 43)
        Me.lblProgress.TabIndex = 15
        '
        'Panel4
        '
        Me.Panel4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel4.BackColor = System.Drawing.Color.SkyBlue
        Me.Panel4.Location = New System.Drawing.Point(1, 454)
        Me.Panel4.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(1004, 12)
        Me.Panel4.TabIndex = 14
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.BackColor = System.Drawing.Color.SkyBlue
        Me.Panel3.Location = New System.Drawing.Point(992, 1)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(13, 465)
        Me.Panel3.TabIndex = 13
        '
        'btnStop
        '
        Me.btnStop.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnStop.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStop.Location = New System.Drawing.Point(879, 89)
        Me.btnStop.Margin = New System.Windows.Forms.Padding(4)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(100, 34)
        Me.btnStop.TabIndex = 12
        Me.btnStop.Text = "Stop"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'lstError
        '
        Me.lstError.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstError.ForeColor = System.Drawing.Color.FromArgb(CType(CType(29, Byte), Integer), CType(CType(29, Byte), Integer), CType(CType(29, Byte), Integer))
        Me.lstError.FormattingEnabled = True
        Me.lstError.HorizontalScrollbar = True
        Me.lstError.ItemHeight = 16
        Me.lstError.Location = New System.Drawing.Point(21, 205)
        Me.lstError.Margin = New System.Windows.Forms.Padding(4)
        Me.lstError.Name = "lstError"
        Me.lstError.Size = New System.Drawing.Size(963, 244)
        Me.lstError.TabIndex = 28
        '
        'lblMainProgress
        '
        Me.lblMainProgress.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblMainProgress.Location = New System.Drawing.Point(22, 131)
        Me.lblMainProgress.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblMainProgress.Name = "lblMainProgress"
        Me.lblMainProgress.Size = New System.Drawing.Size(963, 26)
        Me.lblMainProgress.TabIndex = 29
        Me.lblMainProgress.Text = "Progress Status"
        '
        'frmAllProcess
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1007, 468)
        Me.Controls.Add(Me.chkbAutoProcess)
        Me.Controls.Add(Me.lblMainProgress)
        Me.Controls.Add(Me.lstError)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.grpFolderBrowse)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.btnStop)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmAllProcess"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Score Calculator"
        Me.grpFolderBrowse.ResumeLayout(False)
        Me.grpFolderBrowse.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnStart As Button
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel1 As Panel
    Friend WithEvents btnBrowse As Button
    Friend WithEvents txtFolderpath As TextBox
    Friend WithEvents lblFolderPath As Label
    Friend WithEvents grpFolderBrowse As GroupBox
    Friend WithEvents lblProgress As Label
    Friend WithEvents Panel4 As Panel
    Friend WithEvents Panel3 As Panel
    Friend WithEvents btnStop As Button
    Friend WithEvents folderBrowse As FolderBrowserDialog
    Friend WithEvents lstError As ListBox
    Friend WithEvents lblMainProgress As Label
    Friend WithEvents chkbAutoProcess As CheckBox
End Class
