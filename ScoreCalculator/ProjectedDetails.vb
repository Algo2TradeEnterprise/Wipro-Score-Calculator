Public Class ProjectedDetails
    Public Property Foundation As FoundationDetails
    Public Property ITPi As ITPiDetails
    Public Property PreviousMonthStatusWithFoundation As String
    Public Property PreviousMonthStatusWithoutFoundation As String
    Public ReadOnly Property Progress As Boolean
        Get
            If Me.PreviousMonthStatusWithFoundation IsNot Nothing AndAlso
                (PreviousMonthStatusWithFoundation.ToUpper = "FOUNDATION PENDING" OrElse
                 PreviousMonthStatusWithFoundation.ToUpper = "FOUNDATION COMPLETE") Then
                If Me.Foundation IsNot Nothing AndAlso Me.Foundation.FoundationScore > 30 AndAlso
                    Me.ITPi IsNot Nothing AndAlso Me.ITPi.ITPiScore > 40 Then
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End Get
    End Property
    Public ReadOnly Property ProjectedStatus As String
        Get
            If Me.Progress Then
                If Me.PreviousMonthStatusWithoutFoundation.ToUpper = "FOUNDATION COMPLETE" Then
                    Return "I T Pi"
                Else
                    Return Me.PreviousMonthStatusWithoutFoundation
                End If
            Else
                If Me.PreviousMonthStatusWithFoundation IsNot Nothing Then
                    Return Me.PreviousMonthStatusWithFoundation
                Else
                    If Me.Foundation IsNot Nothing AndAlso Me.Foundation.FoundationScore > 40 Then
                        If Me.ITPi IsNot Nothing AndAlso Me.ITPi.ITPiScore > 65 Then
                            Return "I T Pi"
                        Else
                            Return "Foundation Complete"
                        End If
                    Else
                        Return "Foundation Pending"
                    End If
                End If
            End If
        End Get
    End Property
End Class

Public Class FoundationDetails
    Public Sub New(ByVal towerData As Dictionary(Of String, TowerDetails))
        Me.TowerBucketData = towerData
    End Sub

    Public ReadOnly Property TowerBucketData As Dictionary(Of String, TowerDetails)
    Public ReadOnly Property FoundationScore As Decimal
        Get
            If Me.TowerBucketData IsNot Nothing AndAlso Me.TowerBucketData.Count > 0 Then
                Dim score As Decimal = 0
                For Each runningBucket In Me.TowerBucketData
                    score += runningBucket.Value.CurrentMonthModulatedScore * runningBucket.Value.MappingData.Weightage / 100
                Next
                Return score
            Else
                Return Decimal.MinValue
            End If
        End Get
    End Property
End Class

Public Class ITPiDetails
    Public Sub New(ByVal towerData As Dictionary(Of String, TowerDetails))
        Me.TowerBucketData = towerData
    End Sub

    Public Property TowerBucketData As Dictionary(Of String, TowerDetails)
    Public ReadOnly Property ITPiScore As Decimal
        Get
            If Me.TowerBucketData IsNot Nothing AndAlso Me.TowerBucketData.Count > 0 Then
                Dim score As Decimal = 0
                For Each runningBucket In Me.TowerBucketData
                    score = Math.Max(score, runningBucket.Value.CurrentMonthModulatedScore)
                Next
                Return score
            Else
                Return Decimal.MinValue
            End If
        End Get
    End Property
End Class