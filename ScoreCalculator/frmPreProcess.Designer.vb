<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPreProcess
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPreProcess))
        Me.btnStart = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.txtFolderpath = New System.Windows.Forms.TextBox()
        Me.lblMappingFIle = New System.Windows.Forms.Label()
        Me.grpFolderBrowse = New System.Windows.Forms.GroupBox()
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.folderBrowse = New System.Windows.Forms.FolderBrowserDialog()
        Me.lblError = New System.Windows.Forms.Label()
        Me.btnMappingFileBrowse = New System.Windows.Forms.Button()
        Me.txtMappingFilepath = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.opnMappingFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.grpFolderBrowse.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnStart
        '
        Me.btnStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStart.Location = New System.Drawing.Point(705, 32)
        Me.btnStart.Margin = New System.Windows.Forms.Padding(4)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(100, 34)
        Me.btnStart.TabIndex = 11
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.Black
        Me.Panel2.Location = New System.Drawing.Point(14, 1)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(811, 12)
        Me.Panel2.TabIndex = 10
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Black
        Me.Panel1.Location = New System.Drawing.Point(1, 1)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(13, 469)
        Me.Panel1.TabIndex = 9
        '
        'btnBrowse
        '
        Me.btnBrowse.Location = New System.Drawing.Point(572, 23)
        Me.btnBrowse.Margin = New System.Windows.Forms.Padding(4)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(100, 28)
        Me.btnBrowse.TabIndex = 2
        Me.btnBrowse.Text = "Browse"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'txtFolderpath
        '
        Me.txtFolderpath.Location = New System.Drawing.Point(123, 25)
        Me.txtFolderpath.Margin = New System.Windows.Forms.Padding(4)
        Me.txtFolderpath.Name = "txtFolderpath"
        Me.txtFolderpath.ReadOnly = True
        Me.txtFolderpath.Size = New System.Drawing.Size(444, 22)
        Me.txtFolderpath.TabIndex = 1
        '
        'lblMappingFIle
        '
        Me.lblMappingFIle.AutoSize = True
        Me.lblMappingFIle.Location = New System.Drawing.Point(8, 29)
        Me.lblMappingFIle.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblMappingFIle.Name = "lblMappingFIle"
        Me.lblMappingFIle.Size = New System.Drawing.Size(81, 17)
        Me.lblMappingFIle.TabIndex = 0
        Me.lblMappingFIle.Text = "Folder Path"
        '
        'grpFolderBrowse
        '
        Me.grpFolderBrowse.Controls.Add(Me.btnMappingFileBrowse)
        Me.grpFolderBrowse.Controls.Add(Me.txtMappingFilepath)
        Me.grpFolderBrowse.Controls.Add(Me.Label1)
        Me.grpFolderBrowse.Controls.Add(Me.btnBrowse)
        Me.grpFolderBrowse.Controls.Add(Me.txtFolderpath)
        Me.grpFolderBrowse.Controls.Add(Me.lblMappingFIle)
        Me.grpFolderBrowse.Location = New System.Drawing.Point(21, 17)
        Me.grpFolderBrowse.Margin = New System.Windows.Forms.Padding(4)
        Me.grpFolderBrowse.Name = "grpFolderBrowse"
        Me.grpFolderBrowse.Padding = New System.Windows.Forms.Padding(4)
        Me.grpFolderBrowse.Size = New System.Drawing.Size(676, 106)
        Me.grpFolderBrowse.TabIndex = 8
        Me.grpFolderBrowse.TabStop = False
        Me.grpFolderBrowse.Text = "Choose Folder"
        '
        'lblProgress
        '
        Me.lblProgress.Location = New System.Drawing.Point(21, 127)
        Me.lblProgress.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(786, 71)
        Me.lblProgress.TabIndex = 15
        Me.lblProgress.Text = "Progress Status"
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.Black
        Me.Panel4.Location = New System.Drawing.Point(1, 458)
        Me.Panel4.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(826, 12)
        Me.Panel4.TabIndex = 14
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Black
        Me.Panel3.Location = New System.Drawing.Point(816, 1)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(13, 469)
        Me.Panel3.TabIndex = 13
        '
        'btnStop
        '
        Me.btnStop.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStop.Location = New System.Drawing.Point(705, 73)
        Me.btnStop.Margin = New System.Windows.Forms.Padding(4)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(100, 34)
        Me.btnStop.TabIndex = 12
        Me.btnStop.Text = "Stop"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'lblError
        '
        Me.lblError.Location = New System.Drawing.Point(22, 198)
        Me.lblError.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblError.Name = "lblError"
        Me.lblError.Size = New System.Drawing.Size(786, 256)
        Me.lblError.TabIndex = 16
        Me.lblError.Text = "Error Status"
        '
        'btnMappingFileBrowse
        '
        Me.btnMappingFileBrowse.Location = New System.Drawing.Point(570, 63)
        Me.btnMappingFileBrowse.Margin = New System.Windows.Forms.Padding(4)
        Me.btnMappingFileBrowse.Name = "btnMappingFileBrowse"
        Me.btnMappingFileBrowse.Size = New System.Drawing.Size(100, 28)
        Me.btnMappingFileBrowse.TabIndex = 5
        Me.btnMappingFileBrowse.Text = "Browse"
        Me.btnMappingFileBrowse.UseVisualStyleBackColor = True
        '
        'txtMappingFilepath
        '
        Me.txtMappingFilepath.Location = New System.Drawing.Point(121, 65)
        Me.txtMappingFilepath.Margin = New System.Windows.Forms.Padding(4)
        Me.txtMappingFilepath.Name = "txtMappingFilepath"
        Me.txtMappingFilepath.ReadOnly = True
        Me.txtMappingFilepath.Size = New System.Drawing.Size(444, 22)
        Me.txtMappingFilepath.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(6, 69)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 17)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Mapping File"
        '
        'opnMappingFileDialog
        '
        '
        'frmPreProcess
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(830, 470)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.lblError)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.grpFolderBrowse)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.btnStop)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmPreProcess"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Zero Score Editer"
        Me.grpFolderBrowse.ResumeLayout(False)
        Me.grpFolderBrowse.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnStart As Button
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel1 As Panel
    Friend WithEvents btnBrowse As Button
    Friend WithEvents txtFolderpath As TextBox
    Friend WithEvents lblMappingFIle As Label
    Friend WithEvents grpFolderBrowse As GroupBox
    Friend WithEvents lblProgress As Label
    Friend WithEvents Panel4 As Panel
    Friend WithEvents Panel3 As Panel
    Friend WithEvents btnStop As Button
    Friend WithEvents folderBrowse As FolderBrowserDialog
    Friend WithEvents lblError As Label
    Friend WithEvents btnMappingFileBrowse As Button
    Friend WithEvents txtMappingFilepath As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents opnMappingFileDialog As OpenFileDialog
End Class
