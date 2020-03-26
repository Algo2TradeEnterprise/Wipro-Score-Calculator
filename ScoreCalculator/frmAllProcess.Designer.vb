<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmAllProcess
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmAllProcess))
        Me.btnStartWithScoreAdjustment = New System.Windows.Forms.Button()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.txtFolderpath = New System.Windows.Forms.TextBox()
        Me.lblMappingFIle = New System.Windows.Forms.Label()
        Me.grpFolderBrowse = New System.Windows.Forms.GroupBox()
        Me.chkbAutoProcess = New System.Windows.Forms.CheckBox()
        Me.txtITPi = New System.Windows.Forms.TextBox()
        Me.txtFndtnCmplt = New System.Windows.Forms.TextBox()
        Me.lblITPi = New System.Windows.Forms.Label()
        Me.lblFndtnCmplt = New System.Windows.Forms.Label()
        Me.chkbITPi = New System.Windows.Forms.CheckBox()
        Me.chkbFndtnCmplt = New System.Windows.Forms.CheckBox()
        Me.chkbMaxScoreReplacement = New System.Windows.Forms.CheckBox()
        Me.chkbZeroScoreReplacement = New System.Windows.Forms.CheckBox()
        Me.lblProgress = New System.Windows.Forms.Label()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.folderBrowse = New System.Windows.Forms.FolderBrowserDialog()
        Me.lstError = New System.Windows.Forms.ListBox()
        Me.lblMainProgress = New System.Windows.Forms.Label()
        Me.btnStartWithoutScoreAdjustment = New System.Windows.Forms.Button()
        Me.grpFolderBrowse.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnStartWithScoreAdjustment
        '
        Me.btnStartWithScoreAdjustment.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStartWithScoreAdjustment.Location = New System.Drawing.Point(420, 139)
        Me.btnStartWithScoreAdjustment.Margin = New System.Windows.Forms.Padding(4)
        Me.btnStartWithScoreAdjustment.Name = "btnStartWithScoreAdjustment"
        Me.btnStartWithScoreAdjustment.Size = New System.Drawing.Size(207, 34)
        Me.btnStartWithScoreAdjustment.TabIndex = 11
        Me.btnStartWithScoreAdjustment.Text = "Start With Adjustment"
        Me.btnStartWithScoreAdjustment.UseVisualStyleBackColor = True
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.SkyBlue
        Me.Panel2.Location = New System.Drawing.Point(14, 1)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(991, 12)
        Me.Panel2.TabIndex = 10
        '
        'Panel1
        '
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
        Me.grpFolderBrowse.Controls.Add(Me.chkbAutoProcess)
        Me.grpFolderBrowse.Controls.Add(Me.txtITPi)
        Me.grpFolderBrowse.Controls.Add(Me.txtFndtnCmplt)
        Me.grpFolderBrowse.Controls.Add(Me.lblITPi)
        Me.grpFolderBrowse.Controls.Add(Me.lblFndtnCmplt)
        Me.grpFolderBrowse.Controls.Add(Me.chkbITPi)
        Me.grpFolderBrowse.Controls.Add(Me.chkbFndtnCmplt)
        Me.grpFolderBrowse.Controls.Add(Me.chkbMaxScoreReplacement)
        Me.grpFolderBrowse.Controls.Add(Me.chkbZeroScoreReplacement)
        Me.grpFolderBrowse.Controls.Add(Me.btnBrowse)
        Me.grpFolderBrowse.Controls.Add(Me.txtFolderpath)
        Me.grpFolderBrowse.Controls.Add(Me.lblMappingFIle)
        Me.grpFolderBrowse.Location = New System.Drawing.Point(21, 17)
        Me.grpFolderBrowse.Margin = New System.Windows.Forms.Padding(4)
        Me.grpFolderBrowse.Name = "grpFolderBrowse"
        Me.grpFolderBrowse.Padding = New System.Windows.Forms.Padding(4)
        Me.grpFolderBrowse.Size = New System.Drawing.Size(963, 119)
        Me.grpFolderBrowse.TabIndex = 8
        Me.grpFolderBrowse.TabStop = False
        Me.grpFolderBrowse.Text = "Settings"
        '
        'chkbAutoProcess
        '
        Me.chkbAutoProcess.AutoSize = True
        Me.chkbAutoProcess.Location = New System.Drawing.Point(746, 91)
        Me.chkbAutoProcess.Name = "chkbAutoProcess"
        Me.chkbAutoProcess.Size = New System.Drawing.Size(166, 21)
        Me.chkbAutoProcess.TabIndex = 11
        Me.chkbAutoProcess.Text = "Auto Process All Step"
        Me.chkbAutoProcess.UseVisualStyleBackColor = True
        '
        'txtITPi
        '
        Me.txtITPi.Location = New System.Drawing.Point(406, 90)
        Me.txtITPi.Name = "txtITPi"
        Me.txtITPi.Size = New System.Drawing.Size(100, 22)
        Me.txtITPi.TabIndex = 10
        '
        'txtFndtnCmplt
        '
        Me.txtFndtnCmplt.Location = New System.Drawing.Point(406, 62)
        Me.txtFndtnCmplt.Name = "txtFndtnCmplt"
        Me.txtFndtnCmplt.Size = New System.Drawing.Size(100, 22)
        Me.txtFndtnCmplt.TabIndex = 9
        '
        'lblITPi
        '
        Me.lblITPi.AutoSize = True
        Me.lblITPi.Location = New System.Drawing.Point(324, 92)
        Me.lblITPi.Name = "lblITPi"
        Me.lblITPi.Size = New System.Drawing.Size(75, 17)
        Me.lblITPi.TabIndex = 8
        Me.lblITPi.Text = "Min Score:"
        '
        'lblFndtnCmplt
        '
        Me.lblFndtnCmplt.AutoSize = True
        Me.lblFndtnCmplt.Location = New System.Drawing.Point(324, 64)
        Me.lblFndtnCmplt.Name = "lblFndtnCmplt"
        Me.lblFndtnCmplt.Size = New System.Drawing.Size(75, 17)
        Me.lblFndtnCmplt.TabIndex = 7
        Me.lblFndtnCmplt.Text = "Min Score:"
        '
        'chkbITPi
        '
        Me.chkbITPi.AutoSize = True
        Me.chkbITPi.Location = New System.Drawing.Point(7, 91)
        Me.chkbITPi.Name = "chkbITPi"
        Me.chkbITPi.Size = New System.Drawing.Size(217, 21)
        Me.chkbITPi.TabIndex = 6
        Me.chkbITPi.Text = "Foundation Complete -> I T Pi"
        Me.chkbITPi.UseVisualStyleBackColor = True
        '
        'chkbFndtnCmplt
        '
        Me.chkbFndtnCmplt.AutoSize = True
        Me.chkbFndtnCmplt.Location = New System.Drawing.Point(7, 63)
        Me.chkbFndtnCmplt.Name = "chkbFndtnCmplt"
        Me.chkbFndtnCmplt.Size = New System.Drawing.Size(312, 21)
        Me.chkbFndtnCmplt.TabIndex = 5
        Me.chkbFndtnCmplt.Text = "Foundation Pending -> Foundation Complete"
        Me.chkbFndtnCmplt.UseVisualStyleBackColor = True
        '
        'chkbMaxScoreReplacement
        '
        Me.chkbMaxScoreReplacement.AutoSize = True
        Me.chkbMaxScoreReplacement.Location = New System.Drawing.Point(537, 91)
        Me.chkbMaxScoreReplacement.Name = "chkbMaxScoreReplacement"
        Me.chkbMaxScoreReplacement.Size = New System.Drawing.Size(183, 21)
        Me.chkbMaxScoreReplacement.TabIndex = 4
        Me.chkbMaxScoreReplacement.Text = "Max Score Replacement"
        Me.chkbMaxScoreReplacement.UseVisualStyleBackColor = True
        '
        'chkbZeroScoreReplacement
        '
        Me.chkbZeroScoreReplacement.AutoSize = True
        Me.chkbZeroScoreReplacement.Location = New System.Drawing.Point(537, 63)
        Me.chkbZeroScoreReplacement.Name = "chkbZeroScoreReplacement"
        Me.chkbZeroScoreReplacement.Size = New System.Drawing.Size(188, 21)
        Me.chkbZeroScoreReplacement.TabIndex = 3
        Me.chkbZeroScoreReplacement.Text = "Zero Score Replacement"
        Me.chkbZeroScoreReplacement.UseVisualStyleBackColor = True
        '
        'lblProgress
        '
        Me.lblProgress.Location = New System.Drawing.Point(21, 206)
        Me.lblProgress.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblProgress.Name = "lblProgress"
        Me.lblProgress.Size = New System.Drawing.Size(963, 43)
        Me.lblProgress.TabIndex = 15
        '
        'Panel4
        '
        Me.Panel4.BackColor = System.Drawing.Color.SkyBlue
        Me.Panel4.Location = New System.Drawing.Point(1, 454)
        Me.Panel4.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(1004, 12)
        Me.Panel4.TabIndex = 14
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.Color.SkyBlue
        Me.Panel3.Location = New System.Drawing.Point(992, 1)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(4)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(13, 465)
        Me.Panel3.TabIndex = 13
        '
        'btnStop
        '
        Me.btnStop.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStop.Location = New System.Drawing.Point(879, 139)
        Me.btnStop.Margin = New System.Windows.Forms.Padding(4)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(100, 34)
        Me.btnStop.TabIndex = 12
        Me.btnStop.Text = "Stop"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'lstError
        '
        Me.lstError.ForeColor = System.Drawing.Color.FromArgb(CType(CType(29, Byte), Integer), CType(CType(29, Byte), Integer), CType(CType(29, Byte), Integer))
        Me.lstError.FormattingEnabled = True
        Me.lstError.ItemHeight = 16
        Me.lstError.Location = New System.Drawing.Point(21, 254)
        Me.lstError.Margin = New System.Windows.Forms.Padding(4)
        Me.lstError.Name = "lstError"
        Me.lstError.Size = New System.Drawing.Size(963, 196)
        Me.lstError.TabIndex = 28
        '
        'lblMainProgress
        '
        Me.lblMainProgress.Location = New System.Drawing.Point(22, 180)
        Me.lblMainProgress.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblMainProgress.Name = "lblMainProgress"
        Me.lblMainProgress.Size = New System.Drawing.Size(963, 26)
        Me.lblMainProgress.TabIndex = 29
        Me.lblMainProgress.Text = "Progress Status"
        '
        'btnStartWithoutScoreAdjustment
        '
        Me.btnStartWithoutScoreAdjustment.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnStartWithoutScoreAdjustment.Location = New System.Drawing.Point(635, 139)
        Me.btnStartWithoutScoreAdjustment.Margin = New System.Windows.Forms.Padding(4)
        Me.btnStartWithoutScoreAdjustment.Name = "btnStartWithoutScoreAdjustment"
        Me.btnStartWithoutScoreAdjustment.Size = New System.Drawing.Size(236, 34)
        Me.btnStartWithoutScoreAdjustment.TabIndex = 30
        Me.btnStartWithoutScoreAdjustment.Text = "Start Without Adjustment"
        Me.btnStartWithoutScoreAdjustment.UseVisualStyleBackColor = True
        '
        'frmPreProcess
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1007, 468)
        Me.Controls.Add(Me.btnStartWithoutScoreAdjustment)
        Me.Controls.Add(Me.lblMainProgress)
        Me.Controls.Add(Me.lstError)
        Me.Controls.Add(Me.Panel4)
        Me.Controls.Add(Me.btnStartWithScoreAdjustment)
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
        Me.Text = "Score Calculator"
        Me.grpFolderBrowse.ResumeLayout(False)
        Me.grpFolderBrowse.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnStartWithScoreAdjustment As Button
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
    Friend WithEvents lstError As ListBox
    Friend WithEvents chkbMaxScoreReplacement As CheckBox
    Friend WithEvents chkbZeroScoreReplacement As CheckBox
    Friend WithEvents txtITPi As TextBox
    Friend WithEvents txtFndtnCmplt As TextBox
    Friend WithEvents lblITPi As Label
    Friend WithEvents lblFndtnCmplt As Label
    Friend WithEvents chkbITPi As CheckBox
    Friend WithEvents chkbFndtnCmplt As CheckBox
    Friend WithEvents lblMainProgress As Label
    Friend WithEvents chkbAutoProcess As CheckBox
    Friend WithEvents btnStartWithoutScoreAdjustment As Button
End Class
