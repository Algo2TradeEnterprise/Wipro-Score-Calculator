<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmImpactChange
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
        Me.grpDAAI = New System.Windows.Forms.GroupBox()
        Me.ucImpactDAAI = New ScoreCalculator.ucImpact()
        Me.grpQET = New System.Windows.Forms.GroupBox()
        Me.ucImpactQET = New ScoreCalculator.ucImpact()
        Me.grpMS = New System.Windows.Forms.GroupBox()
        Me.ucImpactMS = New ScoreCalculator.ucImpact()
        Me.grpDX = New System.Windows.Forms.GroupBox()
        Me.ucImpactDX = New ScoreCalculator.ucImpact()
        Me.grpUNIX = New System.Windows.Forms.GroupBox()
        Me.ucImpactUNIX = New ScoreCalculator.ucImpact()
        Me.grpMF = New System.Windows.Forms.GroupBox()
        Me.ucImpactMF = New ScoreCalculator.ucImpact()
        Me.grpDBI = New System.Windows.Forms.GroupBox()
        Me.ucImpactDBI = New ScoreCalculator.ucImpact()
        Me.btnSubmit = New System.Windows.Forms.Button()
        Me.grpDAAI.SuspendLayout()
        Me.grpQET.SuspendLayout()
        Me.grpMS.SuspendLayout()
        Me.grpDX.SuspendLayout()
        Me.grpUNIX.SuspendLayout()
        Me.grpMF.SuspendLayout()
        Me.grpDBI.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpDAAI
        '
        Me.grpDAAI.Controls.Add(Me.ucImpactDAAI)
        Me.grpDAAI.Location = New System.Drawing.Point(2, 2)
        Me.grpDAAI.Name = "grpDAAI"
        Me.grpDAAI.Size = New System.Drawing.Size(513, 100)
        Me.grpDAAI.TabIndex = 0
        Me.grpDAAI.TabStop = False
        Me.grpDAAI.Text = "DAAI"
        '
        'ucImpactDAAI
        '
        Me.ucImpactDAAI.Location = New System.Drawing.Point(6, 22)
        Me.ucImpactDAAI.Name = "ucImpactDAAI"
        Me.ucImpactDAAI.Size = New System.Drawing.Size(494, 62)
        Me.ucImpactDAAI.TabIndex = 0
        '
        'grpQET
        '
        Me.grpQET.Controls.Add(Me.ucImpactQET)
        Me.grpQET.Location = New System.Drawing.Point(2, 108)
        Me.grpQET.Name = "grpQET"
        Me.grpQET.Size = New System.Drawing.Size(513, 100)
        Me.grpQET.TabIndex = 1
        Me.grpQET.TabStop = False
        Me.grpQET.Text = "QET"
        '
        'ucImpactQET
        '
        Me.ucImpactQET.Location = New System.Drawing.Point(6, 22)
        Me.ucImpactQET.Name = "ucImpactQET"
        Me.ucImpactQET.Size = New System.Drawing.Size(494, 62)
        Me.ucImpactQET.TabIndex = 0
        '
        'grpMS
        '
        Me.grpMS.Controls.Add(Me.ucImpactMS)
        Me.grpMS.Location = New System.Drawing.Point(2, 214)
        Me.grpMS.Name = "grpMS"
        Me.grpMS.Size = New System.Drawing.Size(513, 100)
        Me.grpMS.TabIndex = 2
        Me.grpMS.TabStop = False
        Me.grpMS.Text = "MS"
        '
        'ucImpactMS
        '
        Me.ucImpactMS.Location = New System.Drawing.Point(6, 22)
        Me.ucImpactMS.Name = "ucImpactMS"
        Me.ucImpactMS.Size = New System.Drawing.Size(494, 62)
        Me.ucImpactMS.TabIndex = 0
        '
        'grpDX
        '
        Me.grpDX.Controls.Add(Me.ucImpactDX)
        Me.grpDX.Location = New System.Drawing.Point(2, 320)
        Me.grpDX.Name = "grpDX"
        Me.grpDX.Size = New System.Drawing.Size(513, 100)
        Me.grpDX.TabIndex = 3
        Me.grpDX.TabStop = False
        Me.grpDX.Text = "DX"
        '
        'ucImpactDX
        '
        Me.ucImpactDX.Location = New System.Drawing.Point(6, 22)
        Me.ucImpactDX.Name = "ucImpactDX"
        Me.ucImpactDX.Size = New System.Drawing.Size(494, 62)
        Me.ucImpactDX.TabIndex = 0
        '
        'grpUNIX
        '
        Me.grpUNIX.Controls.Add(Me.ucImpactUNIX)
        Me.grpUNIX.Location = New System.Drawing.Point(521, 214)
        Me.grpUNIX.Name = "grpUNIX"
        Me.grpUNIX.Size = New System.Drawing.Size(510, 100)
        Me.grpUNIX.TabIndex = 6
        Me.grpUNIX.TabStop = False
        Me.grpUNIX.Text = "UNIX"
        '
        'ucImpactUNIX
        '
        Me.ucImpactUNIX.Location = New System.Drawing.Point(6, 22)
        Me.ucImpactUNIX.Name = "ucImpactUNIX"
        Me.ucImpactUNIX.Size = New System.Drawing.Size(494, 62)
        Me.ucImpactUNIX.TabIndex = 0
        '
        'grpMF
        '
        Me.grpMF.Controls.Add(Me.ucImpactMF)
        Me.grpMF.Location = New System.Drawing.Point(521, 108)
        Me.grpMF.Name = "grpMF"
        Me.grpMF.Size = New System.Drawing.Size(510, 100)
        Me.grpMF.TabIndex = 5
        Me.grpMF.TabStop = False
        Me.grpMF.Text = "MF"
        '
        'ucImpactMF
        '
        Me.ucImpactMF.Location = New System.Drawing.Point(6, 22)
        Me.ucImpactMF.Name = "ucImpactMF"
        Me.ucImpactMF.Size = New System.Drawing.Size(494, 62)
        Me.ucImpactMF.TabIndex = 0
        '
        'grpDBI
        '
        Me.grpDBI.Controls.Add(Me.ucImpactDBI)
        Me.grpDBI.Location = New System.Drawing.Point(521, 2)
        Me.grpDBI.Name = "grpDBI"
        Me.grpDBI.Size = New System.Drawing.Size(510, 100)
        Me.grpDBI.TabIndex = 4
        Me.grpDBI.TabStop = False
        Me.grpDBI.Text = "DBI"
        '
        'ucImpactDBI
        '
        Me.ucImpactDBI.Location = New System.Drawing.Point(6, 22)
        Me.ucImpactDBI.Name = "ucImpactDBI"
        Me.ucImpactDBI.Size = New System.Drawing.Size(494, 62)
        Me.ucImpactDBI.TabIndex = 0
        '
        'btnSubmit
        '
        Me.btnSubmit.Location = New System.Drawing.Point(923, 374)
        Me.btnSubmit.Name = "btnSubmit"
        Me.btnSubmit.Size = New System.Drawing.Size(106, 46)
        Me.btnSubmit.TabIndex = 7
        Me.btnSubmit.Text = "Submit"
        Me.btnSubmit.UseVisualStyleBackColor = True
        '
        'frmImpactChange
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1036, 423)
        Me.Controls.Add(Me.btnSubmit)
        Me.Controls.Add(Me.grpUNIX)
        Me.Controls.Add(Me.grpMF)
        Me.Controls.Add(Me.grpDBI)
        Me.Controls.Add(Me.grpDX)
        Me.Controls.Add(Me.grpMS)
        Me.Controls.Add(Me.grpQET)
        Me.Controls.Add(Me.grpDAAI)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmImpactChange"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Impact Change"
        Me.grpDAAI.ResumeLayout(False)
        Me.grpQET.ResumeLayout(False)
        Me.grpMS.ResumeLayout(False)
        Me.grpDX.ResumeLayout(False)
        Me.grpUNIX.ResumeLayout(False)
        Me.grpMF.ResumeLayout(False)
        Me.grpDBI.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents grpDAAI As GroupBox
    Friend WithEvents ucImpactDAAI As ucImpact
    Friend WithEvents grpQET As GroupBox
    Friend WithEvents ucImpactQET As ucImpact
    Friend WithEvents grpMS As GroupBox
    Friend WithEvents ucImpactMS As ucImpact
    Friend WithEvents grpDX As GroupBox
    Friend WithEvents ucImpactDX As ucImpact
    Friend WithEvents grpUNIX As GroupBox
    Friend WithEvents ucImpactUNIX As ucImpact
    Friend WithEvents grpMF As GroupBox
    Friend WithEvents ucImpactMF As ucImpact
    Friend WithEvents grpDBI As GroupBox
    Friend WithEvents ucImpactDBI As ucImpact
    Friend WithEvents btnSubmit As Button
End Class
