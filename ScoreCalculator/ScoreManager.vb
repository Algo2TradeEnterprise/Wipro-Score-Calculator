Imports System.IO
Imports Utilities.DAL
Imports System.Threading
Public Class ScoreManager
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

    Private ReadOnly mappingFileSchema As Dictionary(Of String, String)

    Private ReadOnly _cts As CancellationTokenSource
    Private ReadOnly _cmn As Common
    Private ReadOnly _scoreFile As String
    Private ReadOnly _empFile As String
    Private ReadOnly _mappingFile As String

    Private PiDecreased As Integer = 0
    Private TDecreased As Integer = 0
    Private IDecreased As Integer = 0
    Private TIncreased As Integer = 0
    Private IIncreased As Integer = 0
    Private FoundationIncreased As Integer = 0

    Public Sub New(ByVal canceller As CancellationTokenSource,
                   ByVal scoreFile As String,
                   ByVal empFile As String,
                   ByVal mappingFile As String)
        _cts = canceller
        _scoreFile = scoreFile
        _empFile = empFile
        _mappingFile = mappingFile
        _cmn = New Common(_cts)

        mappingFileSchema = New Dictionary(Of String, String) From
            {{"WFT Practice", "WFT Practice"},
            {"WFT Skill Level", "WFT Skill Level"},
            {"WFT Skill Bucket", "WFT Skill Bucket"},
            {"WFT Weightage", "WFT Weightage"},
            {"WFT Subskills", "WFT Subskills"}}
    End Sub

    Public Async Function ProcessData() As Task
        Await Task.Delay(1).ConfigureAwait(False)
        If File.Exists(_scoreFile) Then
            If File.Exists(_empFile) Then
                OnHeartbeat("Opening score file")
                'id, bucket, subskill, wipro skill, score
                Dim empScoreData As Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, Decimal)))) = Nothing
                Using scorexl As New ExcelHelper(_scoreFile, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                    AddHandler scorexl.Heartbeat, AddressOf OnHeartbeat
                    AddHandler scorexl.WaitingFor, AddressOf OnWaitingFor

                    Dim scoreRawData As Dictionary(Of String, Object(,)) = Nothing
                    Dim allSheets As List(Of String) = scorexl.GetExcelSheetsName()
                    If allSheets IsNot Nothing AndAlso allSheets.Count > 0 Then
                        If allSheets.Contains("Summary") Then Throw New ApplicationException("Summary sheet already exists in score file")
                        If allSheets.Contains("Details") Then Throw New ApplicationException("Details sheet already exists in score file")

                        OnHeartbeat("Reading excel in memory")
                        For Each runningSheet In allSheets
                            scorexl.SetActiveSheet(runningSheet)
                            If scoreRawData Is Nothing Then scoreRawData = New Dictionary(Of String, Object(,))
                            scoreRawData.Add(runningSheet, scorexl.GetExcelInMemory())
                        Next
                    End If

                    scorexl.SetActiveSheet("I T Pi")
                    Dim lastEmpScoreRow As Integer = scorexl.GetLastRow(1)
                    For eachEmpRow As Integer = 12 To lastEmpScoreRow
                        OnHeartbeat(String.Format("Creating score collection {0}/{1}", eachEmpRow - 11, lastEmpScoreRow - 11))
                        Dim empID As String = scorexl.GetData(eachEmpRow, 1)
                        For wftBucketColumn As Integer = 1 To 1000
                            Dim wftBucket As String = scorexl.GetData(10, wftBucketColumn)
                            If wftBucket IsNot Nothing Then
                                If wftBucket.ToUpper = "FOUNDATION LEVEL PROFICIENCY" Then
                                    Exit For
                                Else
                                    Dim nextWFTBucketColumn As Integer = Integer.MinValue
                                    For bucketColumn As Integer = wftBucketColumn + 1 To 1000
                                        Dim bucketName As String = scorexl.GetData(10, bucketColumn)
                                        If bucketName IsNot Nothing AndAlso bucketName <> "" Then
                                            nextWFTBucketColumn = bucketColumn
                                            Exit For
                                        End If
                                    Next
                                    If nextWFTBucketColumn <> Integer.MinValue Then
                                        Dim previousWFTsubskill As String = Nothing
                                        For eachWftSubskill As Integer = wftBucketColumn To nextWFTBucketColumn - 3
                                            Dim wftSubskill As String = scorexl.GetData(11, eachWftSubskill)
                                            Dim wftSubskillSheetName As String = FindSheetName(allSheets, wftSubskill, previousWFTsubskill)
                                            If scoreRawData.ContainsKey(wftSubskillSheetName) Then
                                                Dim wftSubskillScore As Object(,) = scoreRawData(wftSubskillSheetName)
                                                Dim empScoreRow As Integer = _cmn.GetRowOf2DArray(wftSubskillScore, 1, empID, True)
                                                If empScoreRow <> Integer.MinValue Then
                                                    For eachWiproSkillColumn As Integer = 3 To wftSubskillScore.GetLength(1) - 1
                                                        Dim wiproSkill As String = wftSubskillScore(1, eachWiproSkillColumn)
                                                        Dim score As Decimal = wftSubskillScore(empScoreRow, eachWiproSkillColumn)
                                                        If empScoreData Is Nothing Then empScoreData = New Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, Decimal))))
                                                        If empScoreData.ContainsKey(empID) Then
                                                            If empScoreData(empID).ContainsKey(wftBucket) Then
                                                                If empScoreData(empID)(wftBucket).ContainsKey(wftSubskill) Then
                                                                    If Not empScoreData(empID)(wftBucket)(wftSubskill).ContainsKey(wiproSkill) Then
                                                                        empScoreData(empID)(wftBucket)(wftSubskill).Add(wiproSkill, score)
                                                                    End If
                                                                Else
                                                                    empScoreData(empID)(wftBucket).Add(wftSubskill, New Dictionary(Of String, Decimal) From
                                                                                                           {{wiproSkill, score}})
                                                                End If
                                                            Else
                                                                empScoreData(empID).Add(wftBucket, New Dictionary(Of String, Dictionary(Of String, Decimal)) From
                                                                                                {{wftSubskill, New Dictionary(Of String, Decimal) From
                                                                                                {{wiproSkill, score}}}})
                                                            End If
                                                        Else
                                                            empScoreData.Add(empID, New Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, Decimal))) From
                                                                                     {{wftBucket, New Dictionary(Of String, Dictionary(Of String, Decimal)) From
                                                                                     {{wftSubskill, New Dictionary(Of String, Decimal) From
                                                                                     {{wiproSkill, score}}}}}})
                                                        End If
                                                    Next
                                                End If
                                            End If
                                            previousWFTsubskill = wftSubskill
                                        Next
                                    End If
                                End If
                            End If
                        Next
                    Next

                    OnHeartbeat("Opening mapping file")
                    Using xl As New ExcelHelper(_mappingFile, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                        AddHandler xl.Heartbeat, AddressOf OnHeartbeat
                        AddHandler xl.WaitingFor, AddressOf OnWaitingFor
                        Dim xlallSheets As List(Of String) = xl.GetExcelSheetsName()
                        If xlallSheets IsNot Nothing AndAlso xlallSheets.Count > 0 Then
                            Dim skillName As String = Path.GetFileNameWithoutExtension(_scoreFile).Split()(0).Trim
                            For Each runningSheet In xlallSheets
                                _cts.Token.ThrowIfCancellationRequested()
                                If runningSheet.ToUpper.Contains(skillName.ToUpper) Then
                                    xl.SetActiveSheet(runningSheet)
                                    OnHeartbeat(String.Format("Checking schema {0}", _mappingFile))
                                    xl.CheckExcelSchema(mappingFileSchema.Values.ToArray)
                                    Exit For
                                End If
                            Next

                            scorexl.CreateNewSheet("Summary")
                            scorexl.SetActiveSheet("Summary")
                            scorexl.SetData(1, 1, "SKILL")
                            scorexl.SetData(1, 2, "Pi Decreased")
                            scorexl.SetData(1, 3, "T Decreased")
                            scorexl.SetData(1, 4, "I Decreased")
                            scorexl.SetData(1, 5, "T Increased")
                            scorexl.SetData(1, 6, "I Increased")
                            scorexl.SetData(1, 7, "Foundation Increased")
                            Dim summaryRowNumber As Integer = 2

                            scorexl.CreateNewSheet("Details")
                            scorexl.SetActiveSheet("Details")
                            scorexl.SetData(1, 1, "EMP ID")
                            scorexl.SetData(1, 2, "ORIGINAL SCORE")
                            scorexl.SetData(1, 3, "CURRENT SCORE")
                            scorexl.SetData(1, 4, "SKILL")
                            Dim detailsRowNumber As Integer = 2

                            Dim skillBucketColumnNumber As Integer = xl.FindAll(mappingFileSchema("WFT Skill Bucket"), xl.GetNamedRange(1, 256, 1, 256), True).FirstOrDefault.Value
                            Dim skillBucketLastRow As Integer = xl.GetLastRow(skillBucketColumnNumber)
                            OnHeartbeat("Finding Wipro Skill")
                            For skillBucketRow As Integer = 2 To skillBucketLastRow
                                _cts.Token.ThrowIfCancellationRequested()
                                Dim skillBucket As String = xl.GetData(skillBucketRow, skillBucketColumnNumber)
                                If skillBucket IsNot Nothing Then
                                    Dim subSkillColumnNumber As Integer = xl.FindAll(mappingFileSchema("WFT Subskills"), xl.GetNamedRange(1, 256, 1, 256), True).FirstOrDefault.Value
                                    Dim subSkillLastRow As Integer = xl.GetLastRow(subSkillColumnNumber)
                                    For subskillRow As Integer = skillBucketRow To subSkillLastRow
                                        _cts.Token.ThrowIfCancellationRequested()
                                        Dim subskill As String = xl.GetData(subskillRow, subSkillColumnNumber)
                                        If subskill IsNot Nothing Then
                                            Dim relatedBucket As String = Nothing
                                            For bucketRow As Integer = subskillRow To 1 Step -1
                                                _cts.Token.ThrowIfCancellationRequested()
                                                Dim bucket As String = xl.GetData(bucketRow, skillBucketColumnNumber)
                                                If bucket IsNot Nothing Then
                                                    relatedBucket = bucket
                                                    Exit For
                                                End If
                                            Next
                                            If relatedBucket IsNot Nothing AndAlso relatedBucket.ToUpper = skillBucket.ToUpper Then
                                                Dim lastColumn As Integer = xl.GetLastCol(subskillRow)
                                                For wiproSkillColumn As Integer = subSkillColumnNumber + 1 To lastColumn
                                                    _cts.Token.ThrowIfCancellationRequested()
                                                    Dim wiproSkill As String = xl.GetData(subskillRow, wiproSkillColumn)
                                                    If wiproSkill IsNot Nothing Then
                                                        Dim empData As DataTable = Nothing
                                                        Using csv As New CSVHelper(_empFile, ",", _cts)
                                                            empData = csv.GetDataTableFromCSV(1)
                                                        End Using
                                                        If empData IsNot Nothing AndAlso empData.Rows.Count > 0 Then
                                                            For row = 1 To empData.Rows.Count - 1
                                                                OnHeartbeat(String.Format("Checking data {0}/{1} for {2}", row, empData.Rows.Count - 1, wiproSkill))
                                                                Dim practice As String = empData.Rows(row).Item(4)
                                                                Dim empID As String = empData.Rows(row).Item(0)
                                                                Dim originalScore As String = empData.Rows(row).Item(1)
                                                                If practice.ToUpper = skillName.ToUpper Then
                                                                    If originalScore = GetFinalScoreOfAnEmployee(empScoreData(empID), Nothing, Nothing, Nothing) Then
                                                                        Dim scoreWithoutWiproskill As String = GetFinalScoreOfAnEmployee(empScoreData(empID), skillBucket, subskill, wiproSkill)
                                                                        scorexl.SetActiveSheet("Details")
                                                                        scorexl.SetData(detailsRowNumber, 1, empID)
                                                                        scorexl.SetData(detailsRowNumber, 2, originalScore)
                                                                        scorexl.SetData(detailsRowNumber, 3, scoreWithoutWiproskill)
                                                                        scorexl.SetData(detailsRowNumber, 4, wiproSkill)
                                                                        detailsRowNumber += 1

                                                                        If originalScore <> scoreWithoutWiproskill Then
                                                                            If originalScore = "Pi" Then
                                                                                PiDecreased = PiDecreased - 1
                                                                            ElseIf originalScore = "T" Then
                                                                                TDecreased = TDecreased - 1
                                                                            ElseIf originalScore = "I" Then
                                                                                IDecreased = IDecreased - 1
                                                                            End If
                                                                            If scoreWithoutWiproskill = "T" Then
                                                                                TIncreased = TIncreased + 1
                                                                            ElseIf scoreWithoutWiproskill = "I" Then
                                                                                IIncreased = IIncreased + 1
                                                                            ElseIf scoreWithoutWiproskill = "Foundation" Then
                                                                                FoundationIncreased = FoundationIncreased + 1
                                                                            End If
                                                                        End If
                                                                    Else
                                                                        Throw New ApplicationException("Check original and calculated score")
                                                                    End If
                                                                Else
                                                                    scorexl.SetActiveSheet("Details")
                                                                    scorexl.SetData(detailsRowNumber, 1, empID)
                                                                    scorexl.SetData(detailsRowNumber, 2, originalScore)
                                                                    scorexl.SetData(detailsRowNumber, 3, originalScore)
                                                                    scorexl.SetData(detailsRowNumber, 4, practice.ToUpper)
                                                                    detailsRowNumber += 1
                                                                End If
                                                            Next
                                                        End If
                                                        scorexl.SetActiveSheet("Summary")
                                                        scorexl.SetData(summaryRowNumber, 1, wiproSkill)
                                                        scorexl.SetData(summaryRowNumber, 2, PiDecreased)
                                                        scorexl.SetData(summaryRowNumber, 3, TDecreased)
                                                        scorexl.SetData(summaryRowNumber, 4, IDecreased)
                                                        scorexl.SetData(summaryRowNumber, 5, TIncreased)
                                                        scorexl.SetData(summaryRowNumber, 6, IIncreased)
                                                        scorexl.SetData(summaryRowNumber, 7, FoundationIncreased)
                                                        summaryRowNumber += 1

                                                        scorexl.SaveExcel()
                                                    End If
                                                Next
                                            Else
                                                Exit For
                                            End If
                                        End If
                                    Next
                                End If
                            Next
                        End If
                    End Using
                End Using
            Else
                Throw New ApplicationException("Emp file not exists")
            End If
        Else
            Throw New ApplicationException("Score file not exists")
        End If
    End Function

    Private Function GetFinalScoreOfAnEmployee(ByVal scoreData As Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, Decimal))),
                                               ByVal bucket As String, ByVal subskill As String, ByVal wiproNom As String) As String
        Dim ret As String = Nothing
        If scoreData IsNot Nothing AndAlso scoreData.Count > 0 Then
            Dim bucketProficiency As Dictionary(Of String, String) = Nothing
            For Each wftBucket In scoreData.Keys
                Dim subskillsScore As Dictionary(Of String, Decimal) = Nothing
                For Each wftSubskill In scoreData(wftBucket).Keys
                    Dim subskillScore As Decimal = Decimal.MinValue
                    For Each wiproSkill In scoreData(wftBucket)(wftSubskill).Keys
                        If Not (bucket IsNot Nothing AndAlso wftBucket.ToUpper = bucket.ToUpper AndAlso
                            subskill IsNot Nothing AndAlso wftSubskill.ToUpper = subskill.ToUpper AndAlso
                            wiproNom IsNot Nothing AndAlso wiproSkill.ToUpper = wiproNom.ToUpper) Then
                            subskillScore = Math.Max(subskillScore, scoreData(wftBucket)(wftSubskill)(wiproSkill))
                        End If
                    Next
                    If subskillsScore Is Nothing Then subskillsScore = New Dictionary(Of String, Decimal)
                    subskillsScore.Add(wftSubskill, subskillScore)
                Next
                If subskillsScore IsNot Nothing AndAlso subskillsScore.Count > 0 Then
                    Dim bucketScore As Decimal = subskillsScore.Max(Function(x)
                                                                        Return x.Value
                                                                    End Function)
                    If bucketProficiency Is Nothing Then bucketProficiency = New Dictionary(Of String, String)
                    bucketProficiency.Add(wftBucket, GetProficiencyLevel(bucketScore))
                End If
            Next
            If bucketProficiency IsNot Nothing AndAlso bucketProficiency.Count > 0 Then
                Dim l3l4 As IEnumerable(Of KeyValuePair(Of String, String)) = bucketProficiency.Where(Function(x)
                                                                                                          If x.Value = "L4" OrElse x.Value = "L3" Then
                                                                                                              Return True
                                                                                                          Else
                                                                                                              Return False
                                                                                                          End If
                                                                                                      End Function)
                Dim l2l3l4 As IEnumerable(Of KeyValuePair(Of String, String)) = bucketProficiency.Where(Function(x)
                                                                                                            If x.Value = "L4" OrElse x.Value = "L3" OrElse x.Value = "L2" Then
                                                                                                                Return True
                                                                                                            Else
                                                                                                                Return False
                                                                                                            End If
                                                                                                        End Function)

                Dim l3l4Count As Integer = 0
                Dim l2l3l4Count As Integer = 0
                If l3l4 IsNot Nothing AndAlso l3l4.Count > 0 Then
                    l3l4Count = Math.Min(2, l3l4.Count)
                End If
                If l2l3l4 IsNot Nothing AndAlso l2l3l4.Count > 0 Then
                    l2l3l4Count = l2l3l4.Count
                End If
                If l3l4Count >= 2 AndAlso l2l3l4Count >= 4 Then
                    ret = "Pi"
                ElseIf l3l4Count >= 1 AndAlso l2l3l4Count >= 3 Then
                    ret = "T"
                ElseIf l3l4Count >= 1 Then
                    ret = "I"
                Else
                    ret = "Foundation"
                End If
            End If
        End If
        Return ret
    End Function

    Private Function GetProficiencyLevel(ByVal score As Decimal) As String
        Dim ret As String = Nothing
        If score < 40.01 Then
            ret = "L1"
        ElseIf score < 65.01 Then
            ret = "L2"
        ElseIf score < 85.01 Then
            ret = "L3"
        Else
            ret = "L4"
        End If
        Return ret
    End Function

    Private Function FindSheetName(ByVal allSheets As List(Of String), ByVal subskill As String, ByVal previousSubskill As String) As String
        Dim ret As String = subskill
        If subskill.Contains(":") Then
            ret = subskill.Replace(":", " ")
        ElseIf subskill.Contains("\") Then
            ret = subskill.Replace("\", " ")
        ElseIf subskill.Contains("/") Then
            ret = subskill.Replace("/", " ")
        ElseIf subskill.Contains("?") Then
            ret = subskill.Replace("?", " ")
        ElseIf subskill.Contains("*") Then
            ret = subskill.Replace("*", " ")
        ElseIf subskill.Contains("[") Then
            ret = subskill.Replace("[", " ")
        ElseIf subskill.Contains("]") Then
            ret = subskill.Replace("]", " ")
        Else
            ret = subskill
        End If
        If ret.Length > 30 Then ret = ret.Substring(0, 25)

        Dim previousSubskillSheet As String = previousSubskill
        If previousSubskill IsNot Nothing Then
            If previousSubskill.Contains(":") Then
                previousSubskillSheet = previousSubskill.Replace(":", " ")
            ElseIf previousSubskill.Contains("\") Then
                previousSubskillSheet = previousSubskill.Replace("\", " ")
            ElseIf previousSubskill.Contains("/") Then
                previousSubskillSheet = previousSubskill.Replace("/", " ")
            ElseIf previousSubskill.Contains("?") Then
                previousSubskillSheet = previousSubskill.Replace("?", " ")
            ElseIf previousSubskill.Contains("*") Then
                previousSubskillSheet = previousSubskill.Replace("*", " ")
            ElseIf previousSubskill.Contains("[") Then
                previousSubskillSheet = previousSubskill.Replace("[", " ")
            ElseIf previousSubskill.Contains("]") Then
                previousSubskillSheet = previousSubskill.Replace("]", " ")
            Else
                previousSubskillSheet = previousSubskill
            End If
            If previousSubskillSheet.Length > 30 Then previousSubskillSheet = previousSubskillSheet.Substring(0, 25)
        End If

        If allSheets IsNot Nothing Then
            Dim previousSheet As String = Nothing
            For i = allSheets.Count - 1 To 0 Step -1
                Dim runningSheet As String = allSheets(i)
                If runningSheet.Contains(ret) Then
                    If previousSubskillSheet Is Nothing OrElse
                        previousSheet Is Nothing Then
                        ret = runningSheet
                        Exit For
                    Else
                        If previousSheet.Contains(previousSubskillSheet) Then
                            ret = runningSheet
                            Exit For
                        End If
                    End If
                End If
                previousSheet = runningSheet
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
