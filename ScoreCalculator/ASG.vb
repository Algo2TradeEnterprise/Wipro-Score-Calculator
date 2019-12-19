Public Class ASG
    Public EmpID As String
    Public WFTPractice As String
    Public WFTSkillLevel As String
    Public WFTSkillBucket As String
    Public ProjectedProficiency As String

    Private _ProjectedScore As Decimal
    Public ReadOnly Property ProjectedScore As Decimal
        Get
            If ProjectedProficiency IsNot Nothing Then
                If ProjectedProficiency = "L1" Then
                    _ProjectedScore = 0.99
                ElseIf ProjectedProficiency = "L2" Then
                    _ProjectedScore = 40.99
                ElseIf ProjectedProficiency = "L3" Then
                    _ProjectedScore = 65.99
                ElseIf ProjectedProficiency = "L4" Then
                    _ProjectedScore = 85.99
                End If
            End If
            Return _ProjectedScore
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return String.Format("{0}-{1}-{2}-{3}", WFTPractice, WFTSkillLevel, WFTSkillBucket, ProjectedProficiency)
    End Function
End Class
