Public Class TowerDetails
    Public Property MappingData As MappingDetails

    Public Property CurrentMonthRawScore As Decimal
    Public Property CurrentMonthModulatedScore As Decimal

    Public Property PreviousMonthRawScore As Decimal
    Public Property PreviousMonthModulatedScore As Decimal
    Public Property IncrementedSkillList As List(Of String)

    Public Property PreviousMonthWithFoundationStatus As String
    Public Property PreviousMonthWithoutFoundationStatus As String
End Class