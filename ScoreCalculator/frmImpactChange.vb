Public Class frmImpactChange
    Private _practiceData As Dictionary(Of String, PracticeDetails) = Nothing

    Public Sub New(ByRef practiceData As Dictionary(Of String, PracticeDetails))
        InitializeComponent()

        _practiceData = practiceData
    End Sub

    Private Sub frmImpactChange_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If _practiceData IsNot Nothing AndAlso _practiceData.Count > 0 Then
            For Each runningPractice In _practiceData
                If runningPractice.Key = "DAAI" Then
                    Dim preMonthPer As Decimal = Math.Round((runningPractice.Value.PreviousMonthITPiEmployeeCount / runningPractice.Value.PreviousMonthTotalEmployeeCount) * 100, 2)
                    Dim curMonthPer As Decimal = Math.Round((runningPractice.Value.CurrentMonthITPiEmployeeCount / runningPractice.Value.CurrentMonthTotalEmployeeCount) * 100, 2)
                    ucImpactDAAI.lblPreviousMonthITPi.Text = String.Format("Previous Month I T Pi: {0} ({1}%)", runningPractice.Value.PreviousMonthITPiEmployeeCount, preMonthPer)
                    ucImpactDAAI.lblCurrentMonthITPi.Text = String.Format("Current Month I T Pi: {0} ({1}%)", runningPractice.Value.CurrentMonthITPiEmployeeCount, curMonthPer)
                    ucImpactDAAI.lblProjectedImpact.Text = String.Format("Projected Impact: {0}%", curMonthPer - preMonthPer)
                    ucImpactDAAI.txtRequiredImpact.Text = curMonthPer - preMonthPer
                ElseIf runningPractice.Key = "QET" Then
                    Dim preMonthPer As Decimal = Math.Round((runningPractice.Value.PreviousMonthITPiEmployeeCount / runningPractice.Value.PreviousMonthTotalEmployeeCount) * 100, 2)
                    Dim curMonthPer As Decimal = Math.Round((runningPractice.Value.CurrentMonthITPiEmployeeCount / runningPractice.Value.CurrentMonthTotalEmployeeCount) * 100, 2)
                    ucImpactQET.lblPreviousMonthITPi.Text = String.Format("Previous Month I T Pi: {0} ({1}%)", runningPractice.Value.PreviousMonthITPiEmployeeCount, preMonthPer)
                    ucImpactQET.lblCurrentMonthITPi.Text = String.Format("Current Month I T Pi: {0} ({1}%)", runningPractice.Value.CurrentMonthITPiEmployeeCount, curMonthPer)
                    ucImpactQET.lblProjectedImpact.Text = String.Format("Projected Impact: {0}%", curMonthPer - preMonthPer)
                    ucImpactQET.txtRequiredImpact.Text = curMonthPer - preMonthPer
                ElseIf runningPractice.Key = "MS" Then
                    Dim preMonthPer As Decimal = Math.Round((runningPractice.Value.PreviousMonthITPiEmployeeCount / runningPractice.Value.PreviousMonthTotalEmployeeCount) * 100, 2)
                    Dim curMonthPer As Decimal = Math.Round((runningPractice.Value.CurrentMonthITPiEmployeeCount / runningPractice.Value.CurrentMonthTotalEmployeeCount) * 100, 2)
                    ucImpactMS.lblPreviousMonthITPi.Text = String.Format("Previous Month I T Pi: {0} ({1}%)", runningPractice.Value.PreviousMonthITPiEmployeeCount, preMonthPer)
                    ucImpactMS.lblCurrentMonthITPi.Text = String.Format("Current Month I T Pi: {0} ({1}%)", runningPractice.Value.CurrentMonthITPiEmployeeCount, curMonthPer)
                    ucImpactMS.lblProjectedImpact.Text = String.Format("Projected Impact: {0}%", curMonthPer - preMonthPer)
                    ucImpactMS.txtRequiredImpact.Text = curMonthPer - preMonthPer
                ElseIf runningPractice.Key = "DX" Then
                    Dim preMonthPer As Decimal = Math.Round((runningPractice.Value.PreviousMonthITPiEmployeeCount / runningPractice.Value.PreviousMonthTotalEmployeeCount) * 100, 2)
                    Dim curMonthPer As Decimal = Math.Round((runningPractice.Value.CurrentMonthITPiEmployeeCount / runningPractice.Value.CurrentMonthTotalEmployeeCount) * 100, 2)
                    ucImpactDX.lblPreviousMonthITPi.Text = String.Format("Previous Month I T Pi: {0} ({1}%)", runningPractice.Value.PreviousMonthITPiEmployeeCount, preMonthPer)
                    ucImpactDX.lblCurrentMonthITPi.Text = String.Format("Current Month I T Pi: {0} ({1}%)", runningPractice.Value.CurrentMonthITPiEmployeeCount, curMonthPer)
                    ucImpactDX.lblProjectedImpact.Text = String.Format("Projected Impact: {0}%", curMonthPer - preMonthPer)
                    ucImpactDX.txtRequiredImpact.Text = curMonthPer - preMonthPer
                ElseIf runningPractice.Key = "DBI" Then
                    Dim preMonthPer As Decimal = Math.Round((runningPractice.Value.PreviousMonthITPiEmployeeCount / runningPractice.Value.PreviousMonthTotalEmployeeCount) * 100, 2)
                    Dim curMonthPer As Decimal = Math.Round((runningPractice.Value.CurrentMonthITPiEmployeeCount / runningPractice.Value.CurrentMonthTotalEmployeeCount) * 100, 2)
                    ucImpactDBI.lblPreviousMonthITPi.Text = String.Format("Previous Month I T Pi: {0} ({1}%)", runningPractice.Value.PreviousMonthITPiEmployeeCount, preMonthPer)
                    ucImpactDBI.lblCurrentMonthITPi.Text = String.Format("Current Month I T Pi: {0} ({1}%)", runningPractice.Value.CurrentMonthITPiEmployeeCount, curMonthPer)
                    ucImpactDBI.lblProjectedImpact.Text = String.Format("Projected Impact: {0}%", curMonthPer - preMonthPer)
                    ucImpactDBI.txtRequiredImpact.Text = curMonthPer - preMonthPer
                ElseIf runningPractice.Key = "MF" Then
                    Dim preMonthPer As Decimal = Math.Round((runningPractice.Value.PreviousMonthITPiEmployeeCount / runningPractice.Value.PreviousMonthTotalEmployeeCount) * 100, 2)
                    Dim curMonthPer As Decimal = Math.Round((runningPractice.Value.CurrentMonthITPiEmployeeCount / runningPractice.Value.CurrentMonthTotalEmployeeCount) * 100, 2)
                    ucImpactMF.lblPreviousMonthITPi.Text = String.Format("Previous Month I T Pi: {0} ({1}%)", runningPractice.Value.PreviousMonthITPiEmployeeCount, preMonthPer)
                    ucImpactMF.lblCurrentMonthITPi.Text = String.Format("Current Month I T Pi: {0} ({1}%)", runningPractice.Value.CurrentMonthITPiEmployeeCount, curMonthPer)
                    ucImpactMF.lblProjectedImpact.Text = String.Format("Projected Impact: {0}%", curMonthPer - preMonthPer)
                    ucImpactMF.txtRequiredImpact.Text = curMonthPer - preMonthPer
                ElseIf runningPractice.Key = "UNIX" Then
                    Dim preMonthPer As Decimal = Math.Round((runningPractice.Value.PreviousMonthITPiEmployeeCount / runningPractice.Value.PreviousMonthTotalEmployeeCount) * 100, 2)
                    Dim curMonthPer As Decimal = Math.Round((runningPractice.Value.CurrentMonthITPiEmployeeCount / runningPractice.Value.CurrentMonthTotalEmployeeCount) * 100, 2)
                    ucImpactUNIX.lblPreviousMonthITPi.Text = String.Format("Previous Month I T Pi: {0} ({1}%)", runningPractice.Value.PreviousMonthITPiEmployeeCount, preMonthPer)
                    ucImpactUNIX.lblCurrentMonthITPi.Text = String.Format("Current Month I T Pi: {0} ({1}%)", runningPractice.Value.CurrentMonthITPiEmployeeCount, curMonthPer)
                    ucImpactUNIX.lblProjectedImpact.Text = String.Format("Projected Impact: {0}%", curMonthPer - preMonthPer)
                    ucImpactUNIX.txtRequiredImpact.Text = curMonthPer - preMonthPer
                End If
            Next
        End If
    End Sub

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        Dim preTotalEmpCount As Integer = 0
        Dim preITpiEmpCount As Integer = 0
        Dim curTotalEmpCount As Integer = 0
        Dim curITpiEmpCount As Integer = 0
        Dim reqEmpCount As Integer = 0
        If _practiceData IsNot Nothing AndAlso _practiceData.Count > 0 Then
            For Each runningPractice In _practiceData
                preTotalEmpCount += runningPractice.Value.PreviousMonthTotalEmployeeCount
                preITpiEmpCount += runningPractice.Value.PreviousMonthITPiEmployeeCount
                curTotalEmpCount += runningPractice.Value.CurrentMonthTotalEmployeeCount
                curITpiEmpCount += runningPractice.Value.CurrentMonthITPiEmployeeCount
                If runningPractice.Key = "DAAI" Then
                    Dim impactPer As Decimal = runningPractice.Value.PreviousMonthImpactPercentage + Val(ucImpactDAAI.txtRequiredImpact.Text)
                    reqEmpCount += Math.Ceiling(runningPractice.Value.CurrentMonthTotalEmployeeCount * impactPer / 100)
                ElseIf runningPractice.Key = "QET" Then
                    Dim impactPer As Decimal = runningPractice.Value.PreviousMonthImpactPercentage + Val(ucImpactQET.txtRequiredImpact.Text)
                    reqEmpCount += Math.Ceiling(runningPractice.Value.CurrentMonthTotalEmployeeCount * impactPer / 100)
                ElseIf runningPractice.Key = "MS" Then
                    Dim impactPer As Decimal = runningPractice.Value.PreviousMonthImpactPercentage + Val(ucImpactMS.txtRequiredImpact.Text)
                    reqEmpCount += Math.Ceiling(runningPractice.Value.CurrentMonthTotalEmployeeCount * impactPer / 100)
                ElseIf runningPractice.Key = "DBI" Then
                    Dim impactPer As Decimal = runningPractice.Value.PreviousMonthImpactPercentage + Val(ucImpactDBI.txtRequiredImpact.Text)
                    reqEmpCount += Math.Ceiling(runningPractice.Value.CurrentMonthTotalEmployeeCount * impactPer / 100)
                ElseIf runningPractice.Key = "DX" Then
                    Dim impactPer As Decimal = runningPractice.Value.PreviousMonthImpactPercentage + Val(ucImpactDX.txtRequiredImpact.Text)
                    reqEmpCount += Math.Ceiling(runningPractice.Value.CurrentMonthTotalEmployeeCount * impactPer / 100)
                ElseIf runningPractice.Key = "MF" Then
                    Dim impactPer As Decimal = runningPractice.Value.PreviousMonthImpactPercentage + Val(ucImpactMF.txtRequiredImpact.Text)
                    reqEmpCount += Math.Ceiling(runningPractice.Value.CurrentMonthTotalEmployeeCount * impactPer / 100)
                ElseIf runningPractice.Key = "UNIX" Then
                    Dim impactPer As Decimal = runningPractice.Value.PreviousMonthImpactPercentage + Val(ucImpactUNIX.txtRequiredImpact.Text)
                    reqEmpCount += Math.Ceiling(runningPractice.Value.CurrentMonthTotalEmployeeCount * impactPer / 100)
                End If
            Next
        End If

        If MessageBox.Show(String.Format("Previous Month: {0}{1}Current Month Projected: {2}{3}Current Month Future: {4}{5}Do you want to proceed?",
                             Math.Round((preITpiEmpCount / preTotalEmpCount) * 100, 2),
                             vbNewLine,
                             Math.Round((curITpiEmpCount / curTotalEmpCount) * 100, 2),
                             vbNewLine,
                             Math.Round((reqEmpCount / curTotalEmpCount) * 100, 2),
                             vbNewLine), "Impact Change", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = MsgBoxResult.Yes Then
            If _practiceData IsNot Nothing AndAlso _practiceData.Count > 0 Then
                For Each runningPractice In _practiceData
                    If runningPractice.Key = "DAAI" Then
                        runningPractice.Value.RequiredImpactPercentage = Val(ucImpactDAAI.txtRequiredImpact.Text)
                    ElseIf runningPractice.Key = "QET" Then
                        runningPractice.Value.RequiredImpactPercentage = Val(ucImpactQET.txtRequiredImpact.Text)
                    ElseIf runningPractice.Key = "MS" Then
                        runningPractice.Value.RequiredImpactPercentage = Val(ucImpactMS.txtRequiredImpact.Text)
                    ElseIf runningPractice.Key = "DBI" Then
                        runningPractice.Value.RequiredImpactPercentage = Val(ucImpactDBI.txtRequiredImpact.Text)
                    ElseIf runningPractice.Key = "DX" Then
                        runningPractice.Value.RequiredImpactPercentage = Val(ucImpactDX.txtRequiredImpact.Text)
                    ElseIf runningPractice.Key = "MF" Then
                        runningPractice.Value.RequiredImpactPercentage = Val(ucImpactMF.txtRequiredImpact.Text)
                    ElseIf runningPractice.Key = "UNIX" Then
                        runningPractice.Value.RequiredImpactPercentage = Val(ucImpactUNIX.txtRequiredImpact.Text)
                    End If
                Next
            End If

            Me.Close()
        End If
    End Sub
End Class