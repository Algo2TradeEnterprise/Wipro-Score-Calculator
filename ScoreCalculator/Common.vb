Imports System.Threading

Public Class Common
    Private ReadOnly _cts As CancellationTokenSource
    Public Sub New(ByVal canceller As CancellationTokenSource)
        _cts = canceller
    End Sub

    Public Function GetColumnOf2DArray(ByVal array As Object(,), ByVal rowNumber As Integer, ByVal searchData As String) As Integer
        Dim ret As Integer = Integer.MinValue
        If array IsNot Nothing AndAlso searchData IsNot Nothing Then
            For column As Integer = 1 To array.GetLength(1)
                _cts.Token.ThrowIfCancellationRequested()
                If array(rowNumber, column) IsNot Nothing AndAlso
                    array(rowNumber, column).ToString.ToUpper = searchData.ToUpper Then
                    ret = column
                    If ret <> Integer.MinValue Then Exit For
                End If
            Next
        End If
        Return ret
    End Function

    Public Function GetRowOf2DArray(ByVal array As Object(,), ByVal columnNumber As Integer, ByVal searchData As String, Optional ByVal ignoreCase As Boolean = False) As Integer
        Dim ret As Integer = Integer.MinValue
        If array IsNot Nothing AndAlso searchData IsNot Nothing Then
            For row As Integer = 1 To array.GetLength(0)
                _cts.Token.ThrowIfCancellationRequested()
                If ignoreCase Then
                    If array(row, columnNumber) IsNot Nothing AndAlso
                    array(row, columnNumber) = searchData Then
                        ret = row
                        If ret <> Integer.MinValue Then Exit For
                    End If
                Else
                    If array(row, columnNumber) IsNot Nothing AndAlso
                    array(row, columnNumber).ToString.ToUpper = searchData.ToUpper Then
                        ret = row
                        If ret <> Integer.MinValue Then Exit For
                    End If
                End If
            Next
        End If
        Return ret
    End Function

    Public Function GetColumnName(ByVal colNum As Integer) As String
        Dim d As Integer
        Dim m As Integer
        Dim name As String
        d = colNum
        name = ""
        Do While (d > 0)
            _cts.Token.ThrowIfCancellationRequested()
            m = (d - 1) Mod 26
            name = Chr(65 + m) + name
            d = Int((d - m) / 26)
        Loop
        Return name
    End Function
End Class
