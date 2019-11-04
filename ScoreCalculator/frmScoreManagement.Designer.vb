<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmScoreManagement
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmScoreManagement))
        Me.btnStart = New System.Windows.Forms.Button()
        Me.opnScoreFile = New System.Windows.Forms.OpenFileDialog()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.lblError = New System.Windows.Forms.Label()
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.pnlFile = New System.Windows.Forms.Panel()
        Me.btnEmpBrowse = New System.Windows.Forms.Button()
        Me.txtEmpFilename = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnMappingBrowse = New System.Windows.Forms.Button()
        Me.txtMappingFilename = New System.Windows.Forms.TextBox()
        Me.lblMappingFilename = New System.Windows.Forms.Label()
        Me.btnScoreBrowse = New System.Windows.Forms.Button()
        Me.txtScoreFilename = New System.Windows.Forms.TextBox()
        Me.lblScoreFilename = New System.Windows.Forms.Label()
        Me.opnMappingFile = New System.Windows.Forms.OpenFileDialog()
        Me.opnEmployeeFile = New System.Windows.Forms.OpenFileDialog()
        Me.pnlFile.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(629, 35)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(75, 51)
        Me.btnStart.TabIndex = 3
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'opnScoreFile
        '
        '
        'btnStop
        '
        Me.btnStop.Location = New System.Drawing.Point(710, 35)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(75, 51)
        Me.btnStop.TabIndex = 4
        Me.btnStop.Text = "Stop"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'lblError
        '
        Me.lblError.Location = New System.Drawing.Point(2, 210)
        Me.lblError.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblError.Name = "lblError"
        Me.lblError.Size = New System.Drawing.Size(786, 148)
        Me.lblError.TabIndex = 18
        Me.lblError.Text = "Error Status"
        '
        'lblProgress
        '
        Me.lblProgress.Location = New System.Drawing.Point(2, 150)
        Me.lblProgress.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(786, 50)
        Me.lblProgress.TabIndex = 17
        Me.lblProgress.Text = "Progress Status"
        '
        'pnlFile
        '
        Me.pnlFile.Controls.Add(Me.btnEmpBrowse)
        Me.pnlFile.Controls.Add(Me.txtEmpFilename)
        Me.pnlFile.Controls.Add(Me.Label2)
        Me.pnlFile.Controls.Add(Me.btnMappingBrowse)
        Me.pnlFile.Controls.Add(Me.txtMappingFilename)
        Me.pnlFile.Controls.Add(Me.lblMappingFilename)
        Me.pnlFile.Controls.Add(Me.btnScoreBrowse)
        Me.pnlFile.Controls.Add(Me.txtScoreFilename)
        Me.pnlFile.Controls.Add(Me.lblScoreFilename)
        Me.pnlFile.Location = New System.Drawing.Point(3, 1)
        Me.pnlFile.Name = "pnlFile"
        Me.pnlFile.Size = New System.Drawing.Size(611, 103)
        Me.pnlFile.TabIndex = 19
        '
        'btnEmpBrowse
        '
        Me.btnEmpBrowse.Location = New System.Drawing.Point(569, 73)
        Me.btnEmpBrowse.Name = "btnEmpBrowse"
        Me.btnEmpBrowse.Size = New System.Drawing.Size(39, 23)
        Me.btnEmpBrowse.TabIndex = 11
        Me.btnEmpBrowse.Text = "..."
        Me.btnEmpBrowse.UseVisualStyleBackColor = True
        '
        'txtEmpFilename
        '
        Me.txtEmpFilename.Location = New System.Drawing.Point(149, 74)
        Me.txtEmpFilename.Name = "txtEmpFilename"
        Me.txtEmpFilename.Size = New System.Drawing.Size(418, 22)
        Me.txtEmpFilename.TabIndex = 10
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(2, 76)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(118, 17)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Choose Emp File:"
        '
        'btnMappingBrowse
        '
        Me.btnMappingBrowse.Location = New System.Drawing.Point(569, 39)
        Me.btnMappingBrowse.Name = "btnMappingBrowse"
        Me.btnMappingBrowse.Size = New System.Drawing.Size(39, 23)
        Me.btnMappingBrowse.TabIndex = 8
        Me.btnMappingBrowse.Text = "..."
        Me.btnMappingBrowse.UseVisualStyleBackColor = True
        '
        'txtMappingFilename
        '
        Me.txtMappingFilename.Location = New System.Drawing.Point(149, 40)
        Me.txtMappingFilename.Name = "txtMappingFilename"
        Me.txtMappingFilename.Size = New System.Drawing.Size(418, 22)
        Me.txtMappingFilename.TabIndex = 7
        '
        'lblMappingFilename
        '
        Me.lblMappingFilename.AutoSize = True
        Me.lblMappingFilename.Location = New System.Drawing.Point(2, 42)
        Me.lblMappingFilename.Name = "lblMappingFilename"
        Me.lblMappingFilename.Size = New System.Drawing.Size(144, 17)
        Me.lblMappingFilename.TabIndex = 6
        Me.lblMappingFilename.Text = "Choose Mapping File:"
        '
        'btnScoreBrowse
        '
        Me.btnScoreBrowse.Location = New System.Drawing.Point(569, 5)
        Me.btnScoreBrowse.Name = "btnScoreBrowse"
        Me.btnScoreBrowse.Size = New System.Drawing.Size(39, 23)
        Me.btnScoreBrowse.TabIndex = 5
        Me.btnScoreBrowse.Text = "..."
        Me.btnScoreBrowse.UseVisualStyleBackColor = True
        '
        'txtScoreFilename
        '
        Me.txtScoreFilename.Location = New System.Drawing.Point(149, 6)
        Me.txtScoreFilename.Name = "txtScoreFilename"
        Me.txtScoreFilename.Size = New System.Drawing.Size(418, 22)
        Me.txtScoreFilename.TabIndex = 4
        '
        'lblScoreFilename
        '
        Me.lblScoreFilename.AutoSize = True
        Me.lblScoreFilename.Location = New System.Drawing.Point(2, 8)
        Me.lblScoreFilename.Name = "lblScoreFilename"
        Me.lblScoreFilename.Size = New System.Drawing.Size(127, 17)
        Me.lblScoreFilename.TabIndex = 3
        Me.lblScoreFilename.Text = "Choose Score File:"
        '
        'opnMappingFile
        '
        '
        'opnEmployeeFile
        '
        '
        'frmScoreManagement
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 366)
        Me.Controls.Add(Me.pnlFile)
        Me.Controls.Add(Me.lblError)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.btnStop)
        Me.Controls.Add(Me.btnStart)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmScoreManagement"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Score Management"
        Me.pnlFile.ResumeLayout(False)
        Me.pnlFile.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnStart As Button
    Friend WithEvents opnScoreFile As OpenFileDialog
    Friend WithEvents btnStop As Button
    Friend WithEvents lblError As Label
    Friend WithEvents lblProgress As Label
    Friend WithEvents pnlFile As Panel
    Friend WithEvents btnEmpBrowse As Button
    Friend WithEvents txtEmpFilename As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents btnMappingBrowse As Button
    Friend WithEvents txtMappingFilename As TextBox
    Friend WithEvents lblMappingFilename As Label
    Friend WithEvents btnScoreBrowse As Button
    Friend WithEvents txtScoreFilename As TextBox
    Friend WithEvents lblScoreFilename As Label
    Friend WithEvents opnMappingFile As OpenFileDialog
    Friend WithEvents opnEmployeeFile As OpenFileDialog
End Class
