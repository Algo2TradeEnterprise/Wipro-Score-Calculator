<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ucImpact
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.lblPreviousMonthITPi = New System.Windows.Forms.Label()
        Me.lblCurrentMonthITPi = New System.Windows.Forms.Label()
        Me.lblProjectedImpact = New System.Windows.Forms.Label()
        Me.lblRequiredImpact = New System.Windows.Forms.Label()
        Me.txtRequiredImpact = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'lblPreviousMonthITPi
        '
        Me.lblPreviousMonthITPi.AutoSize = True
        Me.lblPreviousMonthITPi.Location = New System.Drawing.Point(6, 6)
        Me.lblPreviousMonthITPi.Name = "lblPreviousMonthITPi"
        Me.lblPreviousMonthITPi.Size = New System.Drawing.Size(199, 17)
        Me.lblPreviousMonthITPi.TabIndex = 0
        Me.lblPreviousMonthITPi.Text = "Previous Month I T Pi Count: 0"
        '
        'lblCurrentMonthITPi
        '
        Me.lblCurrentMonthITPi.AutoSize = True
        Me.lblCurrentMonthITPi.Location = New System.Drawing.Point(6, 35)
        Me.lblCurrentMonthITPi.Name = "lblCurrentMonthITPi"
        Me.lblCurrentMonthITPi.Size = New System.Drawing.Size(191, 17)
        Me.lblCurrentMonthITPi.TabIndex = 1
        Me.lblCurrentMonthITPi.Text = "Current Month I T Pi Count: 0"
        '
        'lblProjectedImpact
        '
        Me.lblProjectedImpact.AutoSize = True
        Me.lblProjectedImpact.Location = New System.Drawing.Point(271, 7)
        Me.lblProjectedImpact.Name = "lblProjectedImpact"
        Me.lblProjectedImpact.Size = New System.Drawing.Size(129, 17)
        Me.lblProjectedImpact.TabIndex = 2
        Me.lblProjectedImpact.Text = "Projected Impact: 0"
        '
        'lblRequiredImpact
        '
        Me.lblRequiredImpact.AutoSize = True
        Me.lblRequiredImpact.Location = New System.Drawing.Point(271, 35)
        Me.lblRequiredImpact.Name = "lblRequiredImpact"
        Me.lblRequiredImpact.Size = New System.Drawing.Size(115, 17)
        Me.lblRequiredImpact.TabIndex = 3
        Me.lblRequiredImpact.Text = "Required Impact:"
        '
        'txtRequiredImpact
        '
        Me.txtRequiredImpact.Location = New System.Drawing.Point(390, 34)
        Me.txtRequiredImpact.Name = "txtRequiredImpact"
        Me.txtRequiredImpact.Size = New System.Drawing.Size(100, 22)
        Me.txtRequiredImpact.TabIndex = 4
        Me.txtRequiredImpact.Text = "0"
        '
        'ucImpact
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.txtRequiredImpact)
        Me.Controls.Add(Me.lblRequiredImpact)
        Me.Controls.Add(Me.lblProjectedImpact)
        Me.Controls.Add(Me.lblCurrentMonthITPi)
        Me.Controls.Add(Me.lblPreviousMonthITPi)
        Me.Name = "ucImpact"
        Me.Size = New System.Drawing.Size(504, 62)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblPreviousMonthITPi As Label
    Friend WithEvents lblCurrentMonthITPi As Label
    Friend WithEvents lblProjectedImpact As Label
    Friend WithEvents lblRequiredImpact As Label
    Friend WithEvents txtRequiredImpact As TextBox
End Class
