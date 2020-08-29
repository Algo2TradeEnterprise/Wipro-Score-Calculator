Public Class TowerDetails
    Public Property MappingData As MappingDetails

    Public Property CurrentMonthRawScore As Decimal
    Public Property CurrentMonthModulatedScore As Decimal

    Public Property PreviousMonthRawScore As Decimal
    Public Property PreviousMonthModulatedScore As Decimal
    Public Property IncrementedSkillList As List(Of String)

    Public ReadOnly Property CurrentMonthRawGreaterThanPreviousMonthRawScore As Boolean
        Get
            Return CurrentMonthRawScore > PreviousMonthRawScore
        End Get
    End Property
    Public ReadOnly Property CurrentMonthRawGreaterThanPreviousMonthModulatedScore As Boolean
        Get
            Return CurrentMonthRawScore > PreviousMonthModulatedScore
        End Get
    End Property
    Public ReadOnly Property CurrentMonthRevisedScore As Decimal
        Get
            If CurrentMonthRawGreaterThanPreviousMonthRawScore Then
                If CurrentMonthRawGreaterThanPreviousMonthModulatedScore Then
                    Return CurrentMonthRawScore
                Else
                    Return Math.Min(100, (PreviousMonthModulatedScore + (PreviousMonthModulatedScore * (CurrentMonthRawScore - PreviousMonthRawScore) / 100)))
                End If
            Else
                Return PreviousMonthModulatedScore
            End If
        End Get
    End Property
    Public Property PreviousMonthWithFoundationStatus As String
    Public Property PreviousMonthWithoutFoundationStatus As String
End Class