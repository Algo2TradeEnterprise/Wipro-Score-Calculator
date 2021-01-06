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

#Region "Private Class"
    Private Class WeightageDetails
        Public Bucket As String
        Public Weightage As Decimal
        Public ColumnNumber As Integer
        Public BestScoreColumnNumber As Integer
    End Class

    Private Class PreviousMonthOutputDetails
        Public WithFoundationStatusData As Dictionary(Of String, String) = Nothing
        Public WithoutFoundationStatusData As Dictionary(Of String, String) = Nothing
        Public ITPiEmployeeCount As Integer = 0
        Public TotalEmployeeCount As Integer = 0
        Public ScoreData As Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, Decimal))) = Nothing
    End Class
#End Region

    Private ReadOnly _cts As CancellationTokenSource
    Private ReadOnly _cmn As Common
    Private ReadOnly _currentMonthRawFiles As List(Of String)
    Private ReadOnly _previousMonthRawFiles As List(Of String)
    Private ReadOnly _previousMonthOutputFiles As List(Of String)
    Private ReadOnly _mappingFile As String
    Private ReadOnly scoreFileSchema As Dictionary(Of String, String)
    Private ReadOnly mappingFileSchema As Dictionary(Of String, String)

    Public Sub New(ByVal canceller As CancellationTokenSource,
                   ByVal currentMonthRawFiles As List(Of String), ByVal previousMonthRawFiles As List(Of String),
                   ByVal previousMonthOutputFiles As List(Of String), ByVal mappingFile As String)
        _cts = canceller
        _currentMonthRawFiles = currentMonthRawFiles
        _previousMonthRawFiles = previousMonthRawFiles
        _previousMonthOutputFiles = previousMonthOutputFiles
        _mappingFile = mappingFile

        _cmn = New Common(_cts)
        scoreFileSchema = New Dictionary(Of String, String) From
                          {{"Emp No", "Emp No"},
                           {"Role", "Role"}}

        mappingFileSchema = New Dictionary(Of String, String) From
           {{"WFT Practice", "WFT Practice"},
           {"WFT Skill Level", "WFT Skill Level"},
           {"WFT Skill Bucket", "WFT Skill Bucket"},
           {"WFT Weightage", "WFT Weightage"},
           {"WFT Subskills", "WFT Subskills"}}
    End Sub

    Public Async Function CheckScoreUpdateAsync() As Task(Of Dictionary(Of String, PracticeDetails))
        Await Task.Delay(1).ConfigureAwait(False)
        Dim ret As Dictionary(Of String, PracticeDetails) = Nothing

        Dim previousMonthITPiEmployeeCount As Dictionary(Of String, Integer) = Nothing
        Dim previousMonthTotalEmployeeCount As Dictionary(Of String, Integer) = Nothing

        Dim rawScoreData As Dictionary(Of String, Dictionary(Of String, List(Of ScoreData))) = Nothing
        Dim employeeScoreData As Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, TowerDetails)))) = Nothing
        Dim mappingData As Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, MappingDetails))) = Await GetMappingDataAsync().ConfigureAwait(False)
        If mappingData IsNot Nothing AndAlso mappingData.Count > 0 Then
            If _currentMonthRawFiles IsNot Nothing And _currentMonthRawFiles.Count > 0 AndAlso
                _previousMonthRawFiles IsNot Nothing AndAlso _previousMonthRawFiles.Count > 0 AndAlso
                _previousMonthOutputFiles IsNot Nothing AndAlso _previousMonthOutputFiles.Count > 0 Then
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
                    Dim previousMonthOutputFile As String = Nothing
                    For Each preMonthRunningOutputFile In _previousMonthOutputFiles
                        _cts.Token.ThrowIfCancellationRequested()
                        Dim outputPractice As String = Path.GetFileName(preMonthRunningOutputFile).Split(" ")(0)
                        If outputPractice.ToUpper.Trim = practice.Trim.ToUpper Then
                            previousMonthOutputFile = preMonthRunningOutputFile
                            Exit For
                        End If
                    Next

                    Dim towerBucketData As Dictionary(Of String, MappingDetails) = Nothing
                    If mappingData.ContainsKey(practice.ToUpper) Then
                        For Each runningData In mappingData(practice)
                            For Each innerRunningData In runningData.Value
                                If towerBucketData Is Nothing Then towerBucketData = New Dictionary(Of String, MappingDetails)
                                towerBucketData.Add(innerRunningData.Key, innerRunningData.Value)
                            Next
                        Next
                    End If

                    If previousMonthFile IsNot Nothing AndAlso previousMonthOutputFile IsNot Nothing AndAlso towerBucketData IsNot Nothing AndAlso towerBucketData.Count > 0 Then
                        Dim withFoundationStatus As Dictionary(Of String, String) = Nothing
                        Dim withoutFoundationStatus As Dictionary(Of String, String) = Nothing
                        Dim previousMonthModulatedScore As Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, Decimal))) = Nothing
                        Dim previousMonthOutput As PreviousMonthOutputDetails = Await GetPreviousMonthModulatedScoreAsync(previousMonthOutputFile).ConfigureAwait(False)
                        If previousMonthOutput IsNot Nothing Then
                            withFoundationStatus = previousMonthOutput.WithFoundationStatusData
                            withoutFoundationStatus = previousMonthOutput.WithoutFoundationStatusData
                            previousMonthModulatedScore = previousMonthOutput.ScoreData

                            If previousMonthITPiEmployeeCount Is Nothing Then previousMonthITPiEmployeeCount = New Dictionary(Of String, Integer)
                            previousMonthITPiEmployeeCount.Add(practice.Trim.ToUpper, previousMonthOutput.ITPiEmployeeCount)
                            If previousMonthTotalEmployeeCount Is Nothing Then previousMonthTotalEmployeeCount = New Dictionary(Of String, Integer)
                            previousMonthTotalEmployeeCount.Add(practice.Trim.ToUpper, previousMonthOutput.TotalEmployeeCount)
                        End If
                        If previousMonthModulatedScore IsNot Nothing AndAlso previousMonthModulatedScore.Count > 0 AndAlso
                            withFoundationStatus IsNot Nothing AndAlso withFoundationStatus.Count > 0 AndAlso
                            withoutFoundationStatus IsNot Nothing AndAlso withoutFoundationStatus.Count > 0 Then
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
                                                            Dim scoreDetails As ScoreData = New ScoreData With {.EmpID = empID.ToUpper,
                                                                                                                 .Practice = practice.ToUpper,
                                                                                                                 .SkillName = columnName.ToUpper,
                                                                                                                 .CurrentMonthScore = Val(currentScore.Trim),
                                                                                                                 .PreviousMonthScore = Val(previousScore.Trim)}

                                                            If rawScoreData Is Nothing Then rawScoreData = New Dictionary(Of String, Dictionary(Of String, List(Of ScoreData)))
                                                            If Not rawScoreData.ContainsKey(scoreDetails.Practice) Then
                                                                rawScoreData.Add(scoreDetails.Practice, New Dictionary(Of String, List(Of ScoreData)))
                                                            End If
                                                            If Not rawScoreData(scoreDetails.Practice).ContainsKey(scoreDetails.EmpID) Then
                                                                rawScoreData(scoreDetails.Practice).Add(scoreDetails.EmpID, New List(Of ScoreData))
                                                            End If
                                                            rawScoreData(scoreDetails.Practice)(scoreDetails.EmpID).Add(scoreDetails)
                                                        End If
                                                    End If
                                                End If
                                            Next
                                        End If

                                        For Each runningTower In towerBucketData
                                            _cts.Token.ThrowIfCancellationRequested()
                                            Dim currentMonthRawMaxScore As Decimal = Decimal.MinValue
                                            Dim previousMonthRawMaxScore As Decimal = Decimal.MinValue
                                            Dim previousMonthModulatedMaxScore As Decimal = Decimal.MinValue
                                            Dim previousMonthWithFoundation As String = Nothing
                                            Dim previousMonthWithoutFoundation As String = Nothing
                                            Dim incrementedSkillList As List(Of String) = Nothing

                                            Dim wiproSkillColumnList As List(Of String) = runningTower.Value.WiproSkillColumnList
                                            If wiproSkillColumnList IsNot Nothing AndAlso wiproSkillColumnList.Count > 0 Then
                                                For Each runningColumn In wiproSkillColumnList
                                                    _cts.Token.ThrowIfCancellationRequested()
                                                    Dim currentScoreColumn As Integer = _cmn.GetColumnOf2DArray(currentMonthScoreData, 1, runningColumn)
                                                    Dim currentScore As String = Nothing
                                                    If currentScoreColumn <> Integer.MinValue Then
                                                        currentScore = currentMonthScoreData(rowCounter, currentScoreColumn)
                                                        If IsNumeric(currentScore.Trim) Then
                                                            currentMonthRawMaxScore = Math.Max(currentMonthRawMaxScore, Val(currentScore.Trim))
                                                        End If
                                                    End If

                                                    If previousScoreRow <> Integer.MinValue Then
                                                        Dim previousScoreColumn As Integer = _cmn.GetColumnOf2DArray(previousMonthScoreData, 1, runningColumn)
                                                        Dim previousScore As String = Nothing
                                                        If previousScoreColumn <> Integer.MinValue Then
                                                            previousScore = previousMonthScoreData(previousScoreRow, previousScoreColumn)
                                                            If IsNumeric(previousScore.Trim) Then
                                                                previousMonthRawMaxScore = Math.Max(previousMonthRawMaxScore, Val(previousScore.Trim))

                                                                If Val(currentScore) > Val(previousScore) Then
                                                                    If incrementedSkillList Is Nothing Then incrementedSkillList = New List(Of String)
                                                                    incrementedSkillList.Add(runningColumn)
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Next
                                            End If

                                            If previousMonthModulatedScore.ContainsKey(empID) AndAlso
                                                previousMonthModulatedScore(empID).ContainsKey(runningTower.Value.SkillLevel.Trim.ToUpper) AndAlso
                                                previousMonthModulatedScore(empID)(runningTower.Value.SkillLevel.Trim.ToUpper).ContainsKey(runningTower.Value.SkillBucket.ToUpper) Then
                                                previousMonthModulatedMaxScore = previousMonthModulatedScore(empID)(runningTower.Value.SkillLevel.Trim.ToUpper)(runningTower.Value.SkillBucket.ToUpper)
                                            End If

                                            If withFoundationStatus.ContainsKey(empID) Then
                                                previousMonthWithFoundation = withFoundationStatus(empID)
                                            End If
                                            If withoutFoundationStatus.ContainsKey(empID) Then
                                                previousMonthWithoutFoundation = withoutFoundationStatus(empID)
                                            End If

                                            Dim towerData As TowerDetails = New TowerDetails With {
                                                .MappingData = runningTower.Value,
                                                .CurrentMonthRawScore = currentMonthRawMaxScore,
                                                .PreviousMonthRawScore = previousMonthRawMaxScore,
                                                .PreviousMonthModulatedScore = previousMonthModulatedMaxScore,
                                                .PreviousMonthWithFoundationStatus = previousMonthWithFoundation,
                                                .PreviousMonthWithoutFoundationStatus = previousMonthWithoutFoundation,
                                                .IncrementedSkillList = incrementedSkillList
                                            }

                                            If employeeScoreData Is Nothing Then employeeScoreData = New Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, TowerDetails))))
                                            If Not employeeScoreData.ContainsKey(runningTower.Value.Practice.Trim.ToUpper) Then
                                                employeeScoreData.Add(runningTower.Value.Practice.Trim.ToUpper, New Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, TowerDetails))))
                                            End If
                                            If Not employeeScoreData(runningTower.Value.Practice.Trim.ToUpper).ContainsKey(empID.Trim.ToUpper) Then
                                                employeeScoreData(runningTower.Value.Practice.Trim.ToUpper).Add(empID.Trim.ToUpper, New Dictionary(Of String, Dictionary(Of String, TowerDetails)))
                                            End If
                                            If Not employeeScoreData(runningTower.Value.Practice.Trim.ToUpper)(empID.Trim.ToUpper).ContainsKey(runningTower.Value.SkillLevel.Trim.ToUpper) Then
                                                employeeScoreData(runningTower.Value.Practice.Trim.ToUpper)(empID.Trim.ToUpper).Add(runningTower.Value.SkillLevel.Trim.ToUpper, New Dictionary(Of String, TowerDetails))
                                            End If

                                            employeeScoreData(runningTower.Value.Practice.Trim.ToUpper)(empID.Trim.ToUpper)(runningTower.Value.SkillLevel.Trim.ToUpper).Add(runningTower.Value.SkillBucket.Trim.ToUpper, towerData)
                                        Next
                                    End If
                                Next
                            End If
                        End If
                    End If
                Next
            End If
        End If
        If employeeScoreData IsNot Nothing AndAlso employeeScoreData.Count > 0 Then
            Dim practiceCtr As Integer = 0
            For Each runningPractice In employeeScoreData.Keys
                practiceCtr += 1
                _cts.Token.ThrowIfCancellationRequested()
                Dim totalEmpCount As Integer = 0
                Dim itPiEmpCount As Integer = 0
                For Each runningEmp In employeeScoreData(runningPractice).Keys
                    _cts.Token.ThrowIfCancellationRequested()
                    totalEmpCount += 1

                    OnHeartbeat(String.Format("Checking score update {0}/{1}. Practice:{2}/{3}", totalEmpCount, employeeScoreData(runningPractice).Count, practiceCtr, employeeScoreData.Count))

                    Dim projectedScore As ProjectedDetails = New ProjectedDetails
                    projectedScore.Foundation = New FoundationDetails(employeeScoreData(runningPractice)(runningEmp)("FOUNDATION"))
                    projectedScore.ITPi = New ITPiDetails(employeeScoreData(runningPractice)(runningEmp)("I T PI"))
                    projectedScore.PreviousMonthStatusWithFoundation = employeeScoreData(runningPractice)(runningEmp).Values.FirstOrDefault.Values.FirstOrDefault.PreviousMonthWithFoundationStatus
                    projectedScore.PreviousMonthStatusWithoutFoundation = employeeScoreData(runningPractice)(runningEmp).Values.FirstOrDefault.Values.FirstOrDefault.PreviousMonthWithoutFoundationStatus

                    Dim projectedStatus As String = projectedScore.ProjectedStatus
                    If projectedStatus IsNot Nothing AndAlso
                        (projectedStatus.ToUpper = "I" OrElse
                        projectedStatus.ToUpper = "T" OrElse
                        projectedStatus.ToUpper = "PI" OrElse
                        projectedStatus.ToUpper = "I T PI") Then
                        itPiEmpCount += 1
                    End If

                    If ret Is Nothing Then ret = New Dictionary(Of String, PracticeDetails)
                    If Not ret.ContainsKey(runningPractice) Then
                        ret.Add(runningPractice, New PracticeDetails)
                    End If
                    If ret(runningPractice).ProjectedEmployeeData Is Nothing Then ret(runningPractice).ProjectedEmployeeData = New Dictionary(Of String, ProjectedDetails)
                    ret(runningPractice).ProjectedEmployeeData.Add(runningEmp, projectedScore)
                Next

                If rawScoreData IsNot Nothing AndAlso rawScoreData.ContainsKey(runningPractice) Then
                    ret(runningPractice).RawScoreUpdateData = rawScoreData(runningPractice)
                End If
                ret(runningPractice).CurrentMonthITPiEmployeeCount = itPiEmpCount
                ret(runningPractice).CurrentMonthTotalEmployeeCount = totalEmpCount
                If previousMonthITPiEmployeeCount.ContainsKey(runningPractice) Then
                    ret(runningPractice).PreviousMonthITPiEmployeeCount = previousMonthITPiEmployeeCount(runningPractice)
                End If
                If previousMonthTotalEmployeeCount.ContainsKey(runningPractice) Then
                    ret(runningPractice).PreviousMonthTotalEmployeeCount = previousMonthTotalEmployeeCount(runningPractice)
                End If
            Next
        End If

        Return ret
    End Function

    Private Async Function GetPreviousMonthModulatedScoreAsync(ByVal outputFileName As String) As Task(Of PreviousMonthOutputDetails)
        Await Task.Delay(1).ConfigureAwait(False)
        Dim ret As PreviousMonthOutputDetails = New PreviousMonthOutputDetails
        Dim practice As String = Path.GetFileName(outputFileName).Split(" ")(0)
        OnHeartbeat(String.Format("Opening Previous Month Output File for {0}", practice))
        Using outputXL As New ExcelHelper(outputFileName, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
            OnHeartbeat(String.Format("Getting previous month modulated data for {0} ", practice))
            Dim sheetList As List(Of String) = New List(Of String) From {"Foundation", "I T Pi"}
            Dim sheetCtr As Integer = 0
            For Each runningSheet In sheetList
                sheetCtr += 1
                outputXL.SetActiveSheet(runningSheet)

                Dim weightageData As List(Of WeightageDetails) = Nothing
                Dim lastWeightageColumn As Integer = outputXL.GetLastCol(9)
                For columnCtr As Integer = 1 To lastWeightageColumn
                    Dim weightage As String = outputXL.GetData(9, columnCtr)
                    If weightage IsNot Nothing AndAlso weightage <> "" AndAlso IsNumeric(weightage) Then
                        Dim bucket As String = outputXL.GetData(10, columnCtr)
                        Dim bestScoreColumnNumber As Integer = Integer.MinValue
                        For colCtr As Integer = columnCtr To outputXL.GetLastCol(11)
                            Dim columnName As String = outputXL.GetData(11, colCtr)
                            If columnName.ToUpper = "BEST SCORE FROM GROUP" Then
                                bestScoreColumnNumber = colCtr
                                Exit For
                            End If
                        Next

                        Dim wtgDtls As WeightageDetails = New WeightageDetails With {
                                        .Weightage = CDec(weightage) * 100,
                                        .Bucket = bucket,
                                        .ColumnNumber = columnCtr,
                                        .BestScoreColumnNumber = bestScoreColumnNumber
                                    }

                        If weightageData Is Nothing Then weightageData = New List(Of WeightageDetails)
                        weightageData.Add(wtgDtls)
                    End If
                Next

                If weightageData IsNot Nothing AndAlso weightageData.Count > 0 Then
                    Dim lastRow As Integer = outputXL.GetLastRow(1)
                    For rowCtr As Integer = 12 To lastRow
                        OnHeartbeat(String.Format("Getting previous month modulated data for {0} #{1}/{2} #{3}/{4}", practice, rowCtr - 11, lastRow - 11, sheetCtr, sheetList.Count))
                        Dim empID As String = outputXL.GetData(rowCtr, 1)
                        If empID IsNot Nothing AndAlso empID.Trim <> "" Then
                            For Each runningTower In weightageData
                                Dim score As String = outputXL.GetData(rowCtr, runningTower.BestScoreColumnNumber)
                                If score IsNot Nothing AndAlso IsNumeric(score) Then
                                    If ret.ScoreData Is Nothing Then ret.ScoreData = New Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, Decimal)))
                                    If Not ret.ScoreData.ContainsKey(empID.Trim.ToUpper) Then
                                        ret.ScoreData.Add(empID.Trim.ToUpper, New Dictionary(Of String, Dictionary(Of String, Decimal)))
                                    End If
                                    If Not ret.ScoreData(empID.Trim.ToUpper).ContainsKey(runningSheet.Trim.ToUpper) Then
                                        ret.ScoreData(empID.Trim.ToUpper).Add(runningSheet.Trim.ToUpper, New Dictionary(Of String, Decimal))
                                    End If
                                    ret.ScoreData(empID.Trim.ToUpper)(runningSheet.Trim.ToUpper).Add(runningTower.Bucket.Trim.ToUpper, score)
                                End If
                            Next

                            If runningSheet = "I T Pi" Then
                                Dim lastWeightageCol As Integer = weightageData.LastOrDefault.BestScoreColumnNumber
                                Dim withFoundationStatus As String = outputXL.GetData(rowCtr, lastWeightageCol + 8)
                                Dim withoutFoundationStatus As String = outputXL.GetData(rowCtr, lastWeightageCol + 9)
                                If ret.WithFoundationStatusData Is Nothing Then ret.WithFoundationStatusData = New Dictionary(Of String, String)
                                ret.WithFoundationStatusData.Add(empID, withFoundationStatus)
                                If ret.WithoutFoundationStatusData Is Nothing Then ret.WithoutFoundationStatusData = New Dictionary(Of String, String)
                                ret.WithoutFoundationStatusData.Add(empID, withoutFoundationStatus)
                            End If
                        End If
                    Next
                End If
            Next

            outputXL.SetActiveSheet("Progress")
            Dim grandTotalColumnNumber As Integer = 1
            For colNo As Integer = 1 To 8
                If outputXL.GetData(4, colNo).ToString.ToUpper = "GRAND TOTAL" Then
                    grandTotalColumnNumber = colNo
                    Exit For
                End If
            Next
            ret.ITPiEmployeeCount = outputXL.GetData(7, grandTotalColumnNumber) + outputXL.GetData(8, grandTotalColumnNumber) + outputXL.GetData(9, grandTotalColumnNumber)
            ret.TotalEmployeeCount = outputXL.GetData(10, grandTotalColumnNumber)

            OnHeartbeat(String.Format("Closing Previous Month Output File for {0}", practice))
        End Using
        Return ret
    End Function

    Private Async Function GetMappingDataAsync() As Task(Of Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, MappingDetails))))
        Await Task.Delay(1).ConfigureAwait(False)
        Dim ret As Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, MappingDetails))) = Nothing
        Using xl As New ExcelHelper(_mappingFile, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
            AddHandler xl.Heartbeat, AddressOf OnHeartbeat
            AddHandler xl.WaitingFor, AddressOf OnWaitingFor

            Dim allSheets As List(Of String) = xl.GetExcelSheetsName()
            If allSheets IsNot Nothing AndAlso allSheets.Count > 0 Then
                For Each runningSheet In allSheets
                    _cts.Token.ThrowIfCancellationRequested()
                    OnHeartbeat(String.Format("Checking schema {0}", _mappingFile))
                    xl.CheckExcelSchema(mappingFileSchema.Values.ToArray)
                    If xl.SetActiveSheet(runningSheet) Then
                        OnHeartbeat(String.Format("Reading mapping data for {0}", runningSheet))
                        Dim skillRowColumnNumber As Integer = xl.FindAll(mappingFileSchema("WFT Practice"), xl.GetNamedRange(1, 256, 1, 256), True).FirstOrDefault.Value
                        Dim skillLevelColumnNumber As Integer = xl.FindAll(mappingFileSchema("WFT Skill Level"), xl.GetNamedRange(1, 256, 1, 256), True).FirstOrDefault.Value
                        Dim skillBucketColumnNumber As Integer = xl.FindAll(mappingFileSchema("WFT Skill Bucket"), xl.GetNamedRange(1, 256, 1, 256), True).FirstOrDefault.Value
                        Dim weightageColumnNumber As Integer = xl.FindAll(mappingFileSchema("WFT Weightage"), xl.GetNamedRange(1, 256, 1, 256), True).FirstOrDefault.Value
                        Dim subSkillColumnNumber As Integer = xl.FindAll(mappingFileSchema("WFT Subskills"), xl.GetNamedRange(1, 256, 1, 256), True).FirstOrDefault.Value

                        Dim practiceName As String = xl.GetData(2, skillRowColumnNumber)
                        Dim lastRow As Integer = xl.GetLastRow(subSkillColumnNumber)

                        For skillLevelRow As Integer = 2 To lastRow
                            _cts.Token.ThrowIfCancellationRequested()
                            Dim skillLevel As String = xl.GetData(skillLevelRow, skillLevelColumnNumber)
                            If skillLevel IsNot Nothing Then
                                Dim skillBucket As String = xl.GetData(skillLevelRow, skillBucketColumnNumber)
                                Dim weightage As String = xl.GetData(skillLevelRow, weightageColumnNumber)

                                Dim mappingData As MappingDetails = New MappingDetails With {
                                    .Practice = practiceName,
                                    .SkillLevel = skillLevel,
                                    .SkillBucket = skillBucket,
                                    .Weightage = weightage * 100,
                                    .SubSkills = New Dictionary(Of String, List(Of String))
                                }

                                For subskillRow As Integer = skillLevelRow To lastRow
                                    _cts.Token.ThrowIfCancellationRequested()
                                    Dim subskill As String = xl.GetData(subskillRow, subSkillColumnNumber)
                                    If subskill IsNot Nothing Then
                                        Dim wiproSkillList As List(Of String) = Nothing
                                        Dim lastColumn As Integer = xl.GetLastCol(subskillRow)
                                        For wiproSkillColumn As Integer = subSkillColumnNumber + 1 To lastColumn
                                            _cts.Token.ThrowIfCancellationRequested()
                                            Dim wiproSkill As String = xl.GetData(subskillRow, wiproSkillColumn)
                                            If wiproSkill IsNot Nothing Then
                                                If wiproSkillList Is Nothing Then wiproSkillList = New List(Of String)
                                                wiproSkillList.Add(wiproSkill)
                                            End If
                                        Next
                                        mappingData.SubSkills.Add(subskill, wiproSkillList)
                                    End If

                                    Dim nextSkillBucket As String = xl.GetData(subskillRow + 1, skillBucketColumnNumber)
                                    If nextSkillBucket IsNot Nothing AndAlso nextSkillBucket.Trim.ToUpper <> skillBucket.Trim.ToUpper Then
                                        Exit For
                                    End If
                                Next

                                If ret Is Nothing Then ret = New Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, MappingDetails)))
                                If Not ret.ContainsKey(practiceName.Trim.ToUpper) Then
                                    ret.Add(practiceName.Trim.ToUpper, New Dictionary(Of String, Dictionary(Of String, MappingDetails)))
                                End If
                                If Not ret(practiceName.Trim.ToUpper).ContainsKey(skillLevel.Trim.ToUpper) Then
                                    ret(practiceName.Trim.ToUpper).Add(skillLevel.Trim.ToUpper, New Dictionary(Of String, MappingDetails))
                                End If
                                If Not ret(practiceName.Trim.ToUpper)(skillLevel.Trim.ToUpper).ContainsKey(skillBucket.Trim.ToUpper) Then
                                    ret(practiceName.Trim.ToUpper)(skillLevel.Trim.ToUpper).Add(skillBucket.Trim.ToUpper, mappingData)
                                End If
                            End If
                        Next
                    End If
                Next
            End If
        End Using
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