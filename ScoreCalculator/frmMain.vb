Public Class frmMain
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = String.Format("Score Calculator v{0}", My.Application.Info.Version)
    End Sub

    Private Sub CascadeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CascadeToolStripMenuItem.Click
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseAllToolStripMenuItem.Click
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private Sub InProcessToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InProcessToolStripMenuItem.Click
        Dim frmToShow As New frmInProcess

        With frmToShow
            .MdiParent = Me
            .WindowState = FormWindowState.Normal
            .Show()
        End With
    End Sub

    Private Sub PostProcessToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles PostProcessToolStripMenuItem1.Click
        Dim frmToShow As New frmPostProcess

        With frmToShow
            .MdiParent = Me
            .WindowState = FormWindowState.Normal
            .Show()
        End With
    End Sub

    Private Sub ScoreManagerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ScoreManagerToolStripMenuItem.Click
        Dim frmToShow As New frmScoreManagement

        With frmToShow
            .MdiParent = Me
            .WindowState = FormWindowState.Normal
            .Show()
        End With
    End Sub

    Private Sub PreProcessToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PreProcessToolStripMenuItem.Click
        Dim frmToShow As New frmPreProcess

        With frmToShow
            .MdiParent = Me
            .WindowState = FormWindowState.Normal
            .Show()
        End With
    End Sub
End Class