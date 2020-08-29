Public Class PracticeDetails
    Public Property PracticeName As String
    Public Property ProjectedEmployeeData As Dictionary(Of String, ProjectedDetails)
    Public Property RawScoreUpdateData As Dictionary(Of String, List(Of ScoreData))
    Public Property PreviousMonthTotalEmployeeCount As Integer
    Public Property PreviousMonthITPiEmployeeCount As Integer
    Public ReadOnly Property PreviousMonthImpactPercentage As Decimal
        Get
            If Me.PreviousMonthITPiEmployeeCount <> Integer.MinValue AndAlso Me.PreviousMonthTotalEmployeeCount <> Integer.MinValue Then
                Return Math.Round((Me.PreviousMonthITPiEmployeeCount / Me.PreviousMonthTotalEmployeeCount) * 100, 2)
            Else
                Return Decimal.MinValue
            End If
        End Get
    End Property
    Public Property CurrentMonthTotalEmployeeCount As Integer
    Public Property CurrentMonthITPiEmployeeCount As Integer
    Public ReadOnly Property CurrentMonthImpactPercentage As Decimal
        Get
            If Me.CurrentMonthITPiEmployeeCount <> Integer.MinValue AndAlso Me.CurrentMonthTotalEmployeeCount <> Integer.MinValue Then
                Return Math.Round((Me.CurrentMonthITPiEmployeeCount / Me.CurrentMonthTotalEmployeeCount) * 100, 2)
            Else
                Return Decimal.MinValue
            End If
        End Get
    End Property
    Public ReadOnly Property ProjectedImpactPercentage As Decimal
        Get
            If Me.PreviousMonthImpactPercentage <> Decimal.MinValue AndAlso Me.CurrentMonthImpactPercentage <> Decimal.MinValue Then
                Return (Me.CurrentMonthImpactPercentage - Me.PreviousMonthImpactPercentage)
            Else
                Return Decimal.MinValue
            End If
        End Get
    End Property
    Public Property RequiredImpactPercentage As Decimal
End Class
