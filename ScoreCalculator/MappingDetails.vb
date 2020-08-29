Public Class MappingDetails
    Public Property Practice As String
    Public Property SkillLevel As String
    Public Property SkillBucket As String
    Public Property Weightage As Decimal
    Public Property SubSkills As Dictionary(Of String, List(Of String))
    Public ReadOnly Property WiproSkillColumnList As List(Of String)
        Get
            If Me.SubSkills IsNot Nothing AndAlso Me.SubSkills IsNot Nothing Then
                Dim dataSet As List(Of String) = Nothing
                For Each runningSubskill In Me.SubSkills
                    If runningSubskill.Value IsNot Nothing AndAlso runningSubskill.Value.Count > 0 Then
                        For Each runningWiproSkill In runningSubskill.Value
                            If dataSet Is Nothing Then dataSet = New List(Of String)
                            dataSet.Add(runningWiproSkill)
                        Next
                    End If
                Next
                Return dataSet
            Else
                Return Nothing
            End If
        End Get
    End Property
End Class
