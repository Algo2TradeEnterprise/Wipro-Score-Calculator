<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmScoreModifier
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmScoreModifier))
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.folderBrowse = New System.Windows.Forms.FolderBrowserDialog()
        Me.grpFolderBrowse = New System.Windows.Forms.GroupBox()
        Me.btnEmpBrowse = New System.Windows.Forms.Button()
        Me.txtEmpFilename = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.txtFolderpath = New System.Windows.Forms.TextBox()
        Me.lblMappingFIle = New System.Windows.Forms.Label()
        Me.opnEmployeeFile = New System.Windows.Forms.OpenFileDialog()
        Me.lstError = New System.Windows.Forms.ListBox()
        Me.grpFolderBrowse.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblProgress
        '
        Me.lblProgress.Location = New System.Drawing.Point(7, 112)
        Me.lblProgress.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(786, 50)
        Me.lblProgress.TabIndex = 21
        Me.lblProgress.Text = "Progress Status"
        '
        'btnStop
        '
        Me.btnStop.Location = New System.Drawing.Point(715, 26)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(75, 63)
        Me.btnStop.TabIndex = 20
        Me.btnStop.Text = "Stop"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(634, 26)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(75, 63)
        Me.btnStart.TabIndex = 19
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'grpFolderBrowse
        '
        Me.grpFolderBrowse.Controls.Add(Me.btnEmpBrowse)
        Me.grpFolderBrowse.Controls.Add(Me.txtEmpFilename)
        Me.grpFolderBrowse.Controls.Add(Me.Label2)
        Me.grpFolderBrowse.Controls.Add(Me.btnBrowse)
        Me.grpFolderBrowse.Controls.Add(Me.txtFolderpath)
        Me.grpFolderBrowse.Controls.Add(Me.lblMappingFIle)
        Me.grpFolderBrowse.Location = New System.Drawing.Point(10, 4)
        Me.grpFolderBrowse.Margin = New System.Windows.Forms.Padding(4)
        Me.grpFolderBrowse.Name = "grpFolderBrowse"
        Me.grpFolderBrowse.Padding = New System.Windows.Forms.Padding(4)
        Me.grpFolderBrowse.Size = New System.Drawing.Size(617, 100)
        Me.grpFolderBrowse.TabIndex = 26
        Me.grpFolderBrowse.TabStop = False
        '
        'btnEmpBrowse
        '
        Me.btnEmpBrowse.Location = New System.Drawing.Point(503, 61)
        Me.btnEmpBrowse.Name = "btnEmpBrowse"
        Me.btnEmpBrowse.Size = New System.Drawing.Size(100, 23)
        Me.btnEmpBrowse.TabIndex = 14
        Me.btnEmpBrowse.Text = "Browse"
        Me.btnEmpBrowse.UseVisualStyleBackColor = True
        '
        'txtEmpFilename
        '
        Me.txtEmpFilename.Location = New System.Drawing.Point(126, 62)
        Me.txtEmpFilename.Name = "txtEmpFilename"
        Me.txtEmpFilename.Size = New System.Drawing.Size(365, 22)
        Me.txtEmpFilename.TabIndex = 13
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(5, 64)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(114, 17)
        Me.Label2.TabIndex = 12
        Me.Label2.Text = "Choose Emp File"
        '
        'btnBrowse
        '
        Me.btnBrowse.Location = New System.Drawing.Point(503, 23)
        Me.btnBrowse.Margin = New System.Windows.Forms.Padding(4)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(100, 28)
        Me.btnBrowse.TabIndex = 2
        Me.btnBrowse.Text = "Browse"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'txtFolderpath
        '
        Me.txtFolderpath.Location = New System.Drawing.Point(126, 25)
        Me.txtFolderpath.Margin = New System.Windows.Forms.Padding(4)
        Me.txtFolderpath.Name = "txtFolderpath"
        Me.txtFolderpath.Size = New System.Drawing.Size(365, 22)
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
        'opnEmployeeFile
        '
        '
        'lstError
        '
        Me.lstError.ForeColor = System.Drawing.Color.FromArgb(CType(CType(29, Byte), Integer), CType(CType(29, Byte), Integer), CType(CType(29, Byte), Integer))
        Me.lstError.FormattingEnabled = True
        Me.lstError.ItemHeight = 16
        Me.lstError.Location = New System.Drawing.Point(5, 172)
        Me.lstError.Margin = New System.Windows.Forms.Padding(4)
        Me.lstError.Name = "lstError"
        Me.lstError.Size = New System.Drawing.Size(789, 164)
        Me.lstError.TabIndex = 27
        '
        'frmScoreModifier
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 343)
        Me.Controls.Add(Me.lstError)
        Me.Controls.Add(Me.grpFolderBrowse)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.btnStop)
        Me.Controls.Add(Me.btnStart)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmScoreModifier"
        Me.Text = "Score Modifier"
        Me.grpFolderBrowse.ResumeLayout(False)
        Me.grpFolderBrowse.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lblProgress As Label
    Friend WithEvents btnStop As Button
    Friend WithEvents btnStart As Button
    Friend WithEvents folderBrowse As FolderBrowserDialog
    Friend WithEvents grpFolderBrowse As GroupBox
    Friend WithEvents btnBrowse As Button
    Friend WithEvents txtFolderpath As TextBox
    Friend WithEvents lblMappingFIle As Label
    Friend WithEvents btnEmpBrowse As Button
    Friend WithEvents txtEmpFilename As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents opnEmployeeFile As OpenFileDialog
    Friend WithEvents lstError As ListBox
End Class
