<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmInProcess
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmInProcess))
        Me.btnStart = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lblMappingFIle = New System.Windows.Forms.Label()
        Me.txtMappingFilepath = New System.Windows.Forms.TextBox()
        Me.btnMappingFileBrowse = New System.Windows.Forms.Button()
        Me.grpFileBrowse = New System.Windows.Forms.GroupBox()
        Me.btnEmpDataBrowse = New System.Windows.Forms.Button()
        Me.txtEmployeeDataFilepath = New System.Windows.Forms.TextBox()
        Me.lblEmployeeData = New System.Windows.Forms.Label()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.opnMappingFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.opnEmpDataFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.lblError = New System.Windows.Forms.Label()
        Me.grpFileBrowse.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnStart
        '
        Me.btnStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStart.Location = New System.Drawing.Point(703, 31)
        Me.btnStart.Margin = New System.Windows.Forms.Padding(4)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(100, 34)
        Me.btnStart.TabIndex = 3
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.Black
        Me.Panel2.Location = New System.Drawing.Point(1, 0)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(825, 12)
        Me.Panel2.TabIndex = 2
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.Black
        Me.Panel1.Location = New System.Drawing.Point(1, 0)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(13, 334)
        Me.Panel1.TabIndex = 1
        '
        'lblMappingFIle
        '
        Me.lblMappingFIle.AutoSize = True
        Me.lblMappingFIle.Location = New System.Drawing.Point(8, 26)
        Me.lblMappingFIle.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblMappingFIle.Name = "lblMappingFIle"
        Me.lblMappingFIle.Size = New System.Drawing.Size(88, 17)
        Me.lblMappingFIle.TabIndex = 0
        Me.lblMappingFIle.Text = "Mapping File"
        '
        'txtMappingFilepath
        '
        Me.txtMappingFilepath.Location = New System.Drawing.Point(123, 22)
        Me.txtMappingFilepath.Margin = New System.Windows.Forms.Padding(4)
        Me.txtMappingFilepath.Name = "txtMappingFilepath"
        Me.txtMappingFilepath.ReadOnly = True
        Me.txtMappingFilepath.Size = New System.Drawing.Size(444, 22)
        Me.txtMappingFilepath.TabIndex = 1
        '
        'btnMappingFileBrowse
        '
        Me.btnMappingFileBrowse.Location = New System.Drawing.Point(572, 20)
        Me.btnMappingFileBrowse.Margin = New System.Windows.Forms.Padding(4)
        Me.btnMappingFileBrowse.Name = "btnMappingFileBrowse"
        Me.btnMappingFileBrowse.Size = New System.Drawing.Size(100, 28)
        Me.btnMappingFileBrowse.TabIndex = 2
        Me.btnMappingFileBrowse.Text = "Browse"
        Me.btnMappingFileBrowse.UseVisualStyleBackColor = True
        '
        'grpFileBrowse
        '
        Me.grpFileBrowse.Controls.Add(Me.btnEmpDataBrowse)
        Me.grpFileBrowse.Controls.Add(Me.txtEmployeeDataFilepath)
        Me.grpFileBrowse.Controls.Add(Me.lblEmployeeData)
        Me.grpFileBrowse.Controls.Add(Me.btnMappingFileBrowse)
        Me.grpFileBrowse.Controls.Add(Me.txtMappingFilepath)
        Me.grpFileBrowse.Controls.Add(Me.lblMappingFIle)
        Me.grpFileBrowse.Location = New System.Drawing.Point(19, 16)
        Me.grpFileBrowse.Margin = New System.Windows.Forms.Padding(4)
        Me.grpFileBrowse.Name = "grpFileBrowse"
        Me.grpFileBrowse.Padding = New System.Windows.Forms.Padding(4)
        Me.grpFileBrowse.Size = New System.Drawing.Size(676, 112)
        Me.grpFileBrowse.TabIndex = 0
        Me.grpFileBrowse.TabStop = False
        Me.grpFileBrowse.Text = "Choose File"
        '
        'btnEmpDataBrowse
        '
        Me.btnEmpDataBrowse.Location = New System.Drawing.Point(572, 66)
        Me.btnEmpDataBrowse.Margin = New System.Windows.Forms.Padding(4)
        Me.btnEmpDataBrowse.Name = "btnEmpDataBrowse"
        Me.btnEmpDataBrowse.Size = New System.Drawing.Size(100, 28)
        Me.btnEmpDataBrowse.TabIndex = 5
        Me.btnEmpDataBrowse.Text = "Browse"
        Me.btnEmpDataBrowse.UseVisualStyleBackColor = True
        '
        'txtEmployeeDataFilepath
        '
        Me.txtEmployeeDataFilepath.Location = New System.Drawing.Point(123, 68)
        Me.txtEmployeeDataFilepath.Margin = New System.Windows.Forms.Padding(4)
        Me.txtEmployeeDataFilepath.Name = "txtEmployeeDataFilepath"
        Me.txtEmployeeDataFilepath.ReadOnly = True
        Me.txtEmployeeDataFilepath.Size = New System.Drawing.Size(444, 22)
        Me.txtEmployeeDataFilepath.TabIndex = 4
        '
        'lblEmployeeData
        '
        Me.lblEmployeeData.AutoSize = True
        Me.lblEmployeeData.Location = New System.Drawing.Point(8, 72)
        Me.lblEmployeeData.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblEmployeeData.Name = "lblEmployeeData"
        Me.lblEmployeeData.Size = New System.Drawing.Size(104, 17)
        Me.lblEmployeeData.TabIndex = 3
        Me.lblEmployeeData.Text = "Employee Data"
        '
        'btnStop
        '
        Me.btnStop.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStop.Location = New System.Drawing.Point(703, 72)
        Me.btnStop.Margin = New System.Windows.Forms.Padding(4)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(100, 34)
        Me.btnStop.TabIndex = 4
        Me.btnStop.Text = "Stop"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.Black
        Me.Panel3.Location = New System.Drawing.Point(816, 0)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(13, 334)
        Me.Panel3.TabIndex = 5
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.Black
        Me.Panel4.Location = New System.Drawing.Point(1, 322)
        Me.Panel4.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(828, 12)
        Me.Panel4.TabIndex = 6
        '
        'lblProgress
        '
        Me.lblProgress.Location = New System.Drawing.Point(19, 132)
        Me.lblProgress.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(794, 68)
        Me.lblProgress.TabIndex = 7
        Me.lblProgress.Text = "Progress Status"
        '
        'opnMappingFileDialog
        '
        '
        'opnEmpDataFileDialog
        '
        '
        'lblError
        '
        Me.lblError.Location = New System.Drawing.Point(18, 202)
        Me.lblError.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblError.Name = "lblError"
        Me.lblError.Size = New System.Drawing.Size(795, 115)
        Me.lblError.TabIndex = 17
        Me.lblError.Text = "Error Status"
        '
        'frmInProcess
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(830, 335)
        Me.Controls.Add(Me.lblError)
        Me.Controls.Add(Me.lblProgress)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.btnStop)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.grpFileBrowse)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.MaximizeBox = False
        Me.Name = "frmInProcess"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Score Calculator"
        Me.grpFileBrowse.ResumeLayout(False)
        Me.grpFileBrowse.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnStart As Button
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel1 As Panel
    Friend WithEvents lblMappingFIle As Label
    Friend WithEvents txtMappingFilepath As TextBox
    Friend WithEvents btnMappingFileBrowse As Button
    Friend WithEvents grpFileBrowse As GroupBox
    Friend WithEvents btnStop As Button
    Friend WithEvents Panel3 As Panel
    Friend WithEvents Panel4 As Panel
    Friend WithEvents lblProgress As Label
    Friend WithEvents opnMappingFileDialog As OpenFileDialog
    Friend WithEvents btnEmpDataBrowse As Button
    Friend WithEvents txtEmployeeDataFilepath As TextBox
    Friend WithEvents lblEmployeeData As Label
    Friend WithEvents opnEmpDataFileDialog As OpenFileDialog
    Friend WithEvents lblError As Label
End Class
