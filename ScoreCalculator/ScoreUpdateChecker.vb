Imports System.IO
Imports Utilities.DAL
Imports System.Threading
Public Class ScoreUpdateChecker
    Implements IDisposable

#Region "Events/Event handlers"
    Public Event DocumentDownloadComplete()
    Public Event DocumentRetryStatus(ByVal currentTry As Integer, ByVal totalTries As Integer)
    Public Event Heartbeat(ByVal msg As String)
    Public Event HeartbeatError(ByVal msg As String)
    Public Event WaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
    'The below functions are needed to allow the derived classes to raise the above two events
    Protected Overridable Sub OnDocumentDownloadComplete()
        RaiseEvent DocumentDownloadComplete()
    End Sub
    Protected Overridable Sub OnDocumentRetryStatus(ByVal currentTry As Integer, ByVal totalTries As Integer)
        RaiseEvent DocumentRetryStatus(currentTry, totalTries)
    End Sub
    Protected Overridable Sub OnHeartbeat(ByVal msg As String)
        RaiseEvent Heartbeat(msg)
    End Sub
    Protected Overridable Sub OnHeartbeatError(ByVal msg As String)
        RaiseEvent HeartbeatError(msg)
    End Sub
    Protected Overridable Sub OnWaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
        RaiseEvent WaitingFor(elapsedSecs, totalSecs, msg)
    End Sub
#End Region

    Private ReadOnly _cts As CancellationTokenSource
    Private ReadOnly _cmn As Common
    Private ReadOnly _currentMonthRawFiles As List(Of String)
    Private ReadOnly _previousMonthRawFiles As List(Of String)
    Private ReadOnly scoreFileSchema As Dictionary(Of String, String)

    Public Sub New(ByVal canceller As CancellationTokenSource, ByVal currentMonthRawFiles As List(Of String), ByVal previousMonthRawFiles As List(Of String))
        _cts = canceller
        _currentMonthRawFiles = currentMonthRawFiles
        _previousMonthRawFiles = previousMonthRawFiles

        _cmn = New Common(_cts)
        scoreFileSchema = New Dictionary(Of String, String) From
                          {{"Emp No", "Emp No"},
                           {"Role", "Role"}}

    End Sub

    Public Async Function GetScoreUpdateData() As Task(Of Dictionary(Of String, Dictionary(Of String, List(Of ScoreData))))
        Await Task.Delay(1).ConfigureAwait(False)
        Dim ret As Dictionary(Of String, Dictionary(Of String, List(Of ScoreData))) = Nothing
        If _currentMonthRawFiles IsNot Nothing And _currentMonthRawFiles.Count > 0 AndAlso
            _previousMonthRawFiles IsNot Nothing AndAlso _previousMonthRawFiles.Count > 0 Then
            Dim practiceCtr As Integer = 0
            For Each runningFile In _currentMonthRawFiles
                _cts.Token.ThrowIfCancellationRequested()
                practiceCtr += 1
                Dim currentMonthFileName As String = Path.GetFileName(runningFile)
                Dim practice As String = currentMonthFileName.Split("_")(0)
                Dim previousMonthFile As String = Nothing
                For Each preMonthRunningFile In _previousMonthRawFiles
                    _cts.Token.ThrowIfCancellationRequested()
                    Dim previousMonthFileName As String = Path.GetFileName(preMonthRunningFile)
                    If currentMonthFileName.ToUpper <> previousMonthFileName.ToUpper AndAlso
                        currentMonthFileName.Remove(currentMonthFileName.Count - 13).ToUpper = previousMonthFileName.Remove(previousMonthFileName.Count - 13).ToUpper Then
                        previousMonthFile = preMonthRunningFile
                    End If
                Next
                If previousMonthFile IsNot Nothing Then
                    Dim currentMonthScoreData As Object(,) = Nothing
                    Dim previousMonthScoreData As Object(,) = Nothing
                    OnHeartbeat(String.Format("Opening {0}", runningFile))
                    Using xl As New ExcelHelper(runningFile, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                        AddHandler xl.Heartbeat, AddressOf OnHeartbeat
                        AddHandler xl.WaitingFor, AddressOf OnWaitingFor
                        Try
                            OnHeartbeat(String.Format("Checking schema {0}", runningFile))
                            xl.CheckExcelSchema(scoreFileSchema.Values.ToArray)
                            OnHeartbeat(String.Format("Reading {0}", runningFile))
                            currentMonthScoreData = xl.GetExcelInMemory()
                        Catch ex As Exception
                            OnHeartbeatError(ex.Message)
                        End Try
                    End Using
                    OnHeartbeat(String.Format("Opening {0}", previousMonthFile))
                    Using xl As New ExcelHelper(previousMonthFile, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                        AddHandler xl.Heartbeat, AddressOf OnHeartbeat
                        AddHandler xl.WaitingFor, AddressOf OnWaitingFor
                        Try
                            OnHeartbeat(String.Format("Checking schema {0}", previousMonthFile))
                            xl.CheckExcelSchema(scoreFileSchema.Values.ToArray)
                            OnHeartbeat(String.Format("Reading {0}", previousMonthFile))
                            previousMonthScoreData = xl.GetExcelInMemory()
                        Catch ex As Exception
                            OnHeartbeatError(ex.Message)
                        End Try
                    End Using
                    If currentMonthScoreData IsNot Nothing AndAlso previousMonthScoreData IsNot Nothing Then
                        For rowCounter As Integer = 2 To currentMonthScoreData.GetLength(0) - 1
                            _cts.Token.ThrowIfCancellationRequested()
                            OnHeartbeat(String.Format("Checking score update {0}/{1}. Practice:{2}/{3}", rowCounter, currentMonthScoreData.GetLength(0) - 1, practiceCtr, _currentMonthRawFiles.Count))
                            Dim empID As String = currentMonthScoreData(rowCounter, 1)
                            If empID IsNot Nothing Then
                                Dim previousScoreRow As Integer = _cmn.GetRowOf2DArray(previousMonthScoreData, 1, empID, True)
                                If previousScoreRow <> Integer.MinValue Then
                                    For columnCounter As Integer = 3 To currentMonthScoreData.GetLength(1) - 1
                                        _cts.Token.ThrowIfCancellationRequested()
                                        Dim columnName As String = currentMonthScoreData(1, columnCounter)
                                        Dim currentScore As String = currentMonthScoreData(rowCounter, columnCounter)
                                        Dim previousScoreColumn As Integer = _cmn.GetColumnOf2DArray(previousMonthScoreData, 1, columnName)
                                        Dim previousScore As String = Nothing
                                        If previousScoreColumn <> Integer.MinValue Then
                                            previousScore = previousMonthScoreData(previousScoreRow, previousScoreColumn)
                                        End If
                                        If currentScore IsNot Nothing AndAlso currentScore.Trim <> "" AndAlso IsNumeric(currentScore.Trim) Then
                                            If previousScore IsNot Nothing AndAlso previousScore.Trim <> "" AndAlso IsNumeric(previousScore.Trim) Then
                                                If Val(currentScore.Trim) > Val(previousScore.Trim) Then
                                                    Dim scoreDetails As ScoreData = New ScoreData With
                                                        {.EmpID = empID.ToUpper,
                                                         .Practice = practice.ToUpper,
                                                         .SkillName = columnName.ToUpper,
                                                         .CurrentMonthScore = Val(currentScore.Trim),
                                                         .PreviousMonthScore = Val(previousScore.Trim)}

                                                    If ret Is Nothing Then
                                                        ret = New Dictionary(Of String, Dictionary(Of String, List(Of ScoreData)))
                                                    End If
                                                    If Not ret.ContainsKey(scoreDetails.Practice) Then
                                                        ret.Add(scoreDetails.Practice, New Dictionary(Of String, List(Of ScoreData)))
                                                    End If
                                                    If Not ret(scoreDetails.Practice).ContainsKey(scoreDetails.EmpID) Then
                                                        ret(scoreDetails.Practice).Add(scoreDetails.EmpID, New List(Of ScoreData))
                                                    End If
                                                    ret(scoreDetails.Practice)(scoreDetails.EmpID).Add(scoreDetails)
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        End If
        Return ret
    End Function

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class