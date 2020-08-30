Imports System.IO
Imports Utilities.DAL
Imports System.Threading
Public Class PreProcess
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
    Private ReadOnly _inputDirectoryName As String
    Private ReadOnly _outputDirectoryName As String
    Private ReadOnly _mappingFile As String
    Private ReadOnly _availableScoreUpdates As Dictionary(Of String, PracticeDetails) = Nothing
    Private ReadOnly monthList As Dictionary(Of String, String)
    Private ReadOnly scoreFileSchema As Dictionary(Of String, String)
    Private ReadOnly empFileSchema As Dictionary(Of String, String)

    Public Sub New(ByVal canceller As CancellationTokenSource, ByVal directoryName As String, ByVal mappingFile As String,
                   ByVal availableScoreUpdates As Dictionary(Of String, PracticeDetails))
        _cts = canceller
        _inputDirectoryName = directoryName
        _outputDirectoryName = Path.Combine(Directory.GetParent(_inputDirectoryName).FullName, "In Process")
        _mappingFile = mappingFile
        _availableScoreUpdates = availableScoreUpdates

        _cmn = New Common(_cts)
        monthList = New Dictionary(Of String, String) From
                    {{"JAN", "DEC"},
                     {"FEB", "JAN"},
                     {"MAR", "FEB"},
                     {"APR", "MAR"},
                     {"MAY", "APR"},
                     {"JUN", "MAY"},
                     {"JUL", "JUN"},
                     {"AUG", "JUL"},
                     {"SEP", "AUG"},
                     {"OCT", "SEP"},
                     {"NOV", "OCT"},
                     {"DEC", "NOV"}
                    }
        scoreFileSchema = New Dictionary(Of String, String) From
                          {{"Emp No", "Emp No"},
                           {"Role", "Role"}}
        empFileSchema = New Dictionary(Of String, String) From
            {{"Emp No", "EMP_CODE"},
            {"Name", "EMP_NAME"},
            {"Career_Band", "CAREER_BAND"},
            {"Vertical", "SAP_VERTICAL_DESC"},
            {"Normalized Account", "GROUP_CUSTOMER_NAME"},
            {"Email ID", "EMPLOYEE_EMAIL_ID"},
            {"Final Sub Practice", "SAP_SUBPRAC_DESC"},
            {"Project Acquired Skills", "PROJECT_ACQUIRED_SKILL"},
            {"Recent Skills", "RECENT_SKILLS"},
            {"Experience", "EXPERIENCE"},
            {"Applicable for WFT", "Applicable for WFT"},
            {"Exclusion Reason", "Exclusion Reason"},
            {"Status With Foundation", "Status With Foundation"},
            {"Status Without Foundation", "Status Without Foundation"},
            {"Pi Approach", "Pi Approach"},
            {"Last Month Level", "Last Month Level"},
            {"Experience Ok", "Experience Ok?"},
            {"Account Remarks", "Account - Remarks"},
            {"Account Status", "Account - Status"},
            {"Final Account Validation", "Final Account Validation"}}
    End Sub

    Public Async Function ProcessScoreData() As Task
        Await Task.Delay(1).ConfigureAwait(False)
        If Not Directory.Exists(_outputDirectoryName) Then
            Throw New ApplicationException(String.Format("Directory not found. {0}", _outputDirectoryName))
        End If
        Dim practiceWiseMaxScoreData As Dictionary(Of String, Object(,)) = Nothing
        Dim practiceWiseFilename As Dictionary(Of String, String) = Nothing
        Dim skillList As List(Of String) = Nothing
        For Each runningFile In Directory.GetFiles(_inputDirectoryName)
            If runningFile.ToUpper.Contains("BFSI") Then
                Continue For
            End If
            If skillList Is Nothing OrElse Not skillList.Contains(runningFile, StringComparer.OrdinalIgnoreCase) Then
                Dim pairFound As Boolean = False
                For Each runningPairFile In Directory.GetFiles(_inputDirectoryName)
                    If runningPairFile.ToUpper <> runningFile.ToUpper AndAlso
                        runningPairFile.Remove(runningPairFile.Count - 13).ToUpper = runningFile.Remove(runningFile.Count - 13).ToUpper Then
                        pairFound = True
                        If skillList Is Nothing Then skillList = New List(Of String)
                        skillList.Add(runningFile)
                        skillList.Add(runningPairFile)
                        Dim pairFileMonth As String = runningPairFile.Substring(runningPairFile.Count - 13, 3)
                        Dim mainFileMonth As String = runningFile.Substring(runningFile.Count - 13, 3)
                        Dim pairFilePreviousMonth As String = Nothing
                        Dim mainFilePreviousMonth As String = Nothing
                        If monthList.ContainsKey(pairFileMonth.ToUpper) Then pairFilePreviousMonth = monthList(pairFileMonth.ToUpper)
                        If monthList.ContainsKey(mainFileMonth.ToUpper) Then mainFilePreviousMonth = monthList(mainFileMonth.ToUpper)

                        Dim currentFileName As String = Nothing
                        Dim previousFileName As String = Nothing
                        If pairFilePreviousMonth IsNot Nothing AndAlso pairFilePreviousMonth.ToUpper = mainFileMonth.ToUpper Then
                            currentFileName = runningPairFile
                            previousFileName = runningFile
                        ElseIf mainFilePreviousMonth IsNot Nothing AndAlso mainFilePreviousMonth.ToUpper = pairFileMonth.ToUpper Then
                            currentFileName = runningFile
                            previousFileName = runningPairFile
                        End If

                        If currentFileName IsNot Nothing AndAlso previousFileName IsNot Nothing Then
                            Dim currentMonthScoreData As Object(,) = Nothing
                            Dim previousMonthScoreData As Object(,) = Nothing
                            OnHeartbeat(String.Format("Opening {0}", currentFileName))
                            Using xl As New ExcelHelper(currentFileName, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                                AddHandler xl.Heartbeat, AddressOf OnHeartbeat
                                AddHandler xl.WaitingFor, AddressOf OnWaitingFor
                                Try
                                    OnHeartbeat(String.Format("Checking schema {0}", currentFileName))
                                    xl.CheckExcelSchema(scoreFileSchema.Values.ToArray)
                                    OnHeartbeat(String.Format("Reading {0}", currentFileName))
                                    currentMonthScoreData = xl.GetExcelInMemory()
                                Catch ex As Exception
                                    OnHeartbeatError(ex.Message)
                                End Try
                            End Using
                            OnHeartbeat(String.Format("Opening {0}", previousFileName))
                            Using xl As New ExcelHelper(previousFileName, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                                AddHandler xl.Heartbeat, AddressOf OnHeartbeat
                                AddHandler xl.WaitingFor, AddressOf OnWaitingFor
                                Try
                                    OnHeartbeat(String.Format("Checking schema {0}", previousFileName))
                                    xl.CheckExcelSchema(scoreFileSchema.Values.ToArray)
                                    OnHeartbeat(String.Format("Reading {0}", previousFileName))
                                    previousMonthScoreData = xl.GetExcelInMemory()
                                Catch ex As Exception
                                    OnHeartbeatError(ex.Message)
                                End Try
                            End Using
                            If currentMonthScoreData IsNot Nothing AndAlso previousMonthScoreData IsNot Nothing Then
                                Dim practice As String = Path.GetFileName(currentFileName).Split("_")(0)
                                For rowCounter As Integer = 2 To currentMonthScoreData.GetLength(0) - 1
                                    OnHeartbeat(String.Format("Replacing Max score {0}/{1}. Practice:{2}", rowCounter, currentMonthScoreData.GetLength(0) - 1, practice))
                                    _cts.Token.ThrowIfCancellationRequested()
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
                                                If previousScore IsNot Nothing AndAlso IsNumeric(previousScore) Then
                                                    Dim updatedScore As Decimal = Decimal.MinValue
                                                    If _availableScoreUpdates.ContainsKey(practice.Trim.ToUpper) Then
                                                        If _availableScoreUpdates(practice.Trim.ToUpper) IsNot Nothing AndAlso
                                                                _availableScoreUpdates(practice.Trim.ToUpper).RawScoreUpdateData IsNot Nothing AndAlso
                                                                _availableScoreUpdates(practice.Trim.ToUpper).RawScoreUpdateData.ContainsKey(empID.Trim.ToUpper) Then
                                                            Dim scoreDetails As ScoreData =
                                                                    _availableScoreUpdates(practice.Trim.ToUpper).RawScoreUpdateData(empID.Trim.ToUpper).Find(Function(x)
                                                                                                                                                                  Return x.SkillName.ToUpper = columnName.ToUpper
                                                                                                                                                              End Function)
                                                            If scoreDetails IsNot Nothing AndAlso scoreDetails.CurrentMonthScore > scoreDetails.PreviousMonthScore Then
                                                                updatedScore = Math.Min(100, previousScore + (scoreDetails.CurrentMonthScore - scoreDetails.PreviousMonthScore) / 100)
                                                            End If
                                                        End If
                                                    End If

                                                    currentMonthScoreData(rowCounter, columnCounter) = Math.Ceiling(Math.Max(Val(currentScore), Math.Max(updatedScore, Val(previousScore))))
                                                End If
                                            Next
                                        End If
                                    End If
                                Next

                                If _availableScoreUpdates.ContainsKey(practice.Trim.ToUpper) Then
                                    If _availableScoreUpdates(practice.Trim.ToUpper) IsNot Nothing AndAlso
                                        _availableScoreUpdates(practice.Trim.ToUpper).ProjectedEmployeeData IsNot Nothing AndAlso
                                        _availableScoreUpdates(practice.Trim.ToUpper).ProjectedEmployeeData.Count Then
                                        Dim practiceData As PracticeDetails = _availableScoreUpdates(practice.Trim.ToUpper)

                                        Dim totalEmpCount As Integer = 0
                                        Dim itPiEmpCount As Integer = 0
                                        For Each runningEmp In practiceData.ProjectedEmployeeData.Keys
                                            totalEmpCount += 1
                                            OnHeartbeat(String.Format("Updating Current Month modulated score {0}/{1}. Practice:{2}", totalEmpCount, practiceData.ProjectedEmployeeData.Count, practice))
                                            Dim projectedData As ProjectedDetails = practiceData.ProjectedEmployeeData(runningEmp)
                                            If projectedData IsNot Nothing AndAlso projectedData.Foundation IsNot Nothing AndAlso projectedData.Foundation.TowerBucketData IsNot Nothing Then
                                                UpdateCurrentMonthModulatedScore(projectedData.Foundation.TowerBucketData, currentMonthScoreData, runningEmp)
                                            End If
                                            If projectedData IsNot Nothing AndAlso projectedData.ITPi IsNot Nothing AndAlso projectedData.ITPi.TowerBucketData IsNot Nothing Then
                                                UpdateCurrentMonthModulatedScore(projectedData.ITPi.TowerBucketData, currentMonthScoreData, runningEmp)
                                            End If

                                            Dim projectedStatus As String = projectedData.ProjectedStatus
                                            If projectedStatus IsNot Nothing AndAlso
                                                    (projectedStatus.ToUpper = "I" OrElse
                                                    projectedStatus.ToUpper = "T" OrElse
                                                    projectedStatus.ToUpper = "PI" OrElse
                                                    projectedStatus.ToUpper = "I T PI") Then
                                                itPiEmpCount += 1
                                            End If
                                        Next
                                        practiceData.CurrentMonthITPiEmployeeCount = itPiEmpCount
                                    End If
                                End If

                                If practiceWiseMaxScoreData Is Nothing Then practiceWiseMaxScoreData = New Dictionary(Of String, Object(,))
                                practiceWiseMaxScoreData.Add(practice.ToUpper.Trim, currentMonthScoreData)

                                If practiceWiseFilename Is Nothing Then practiceWiseFilename = New Dictionary(Of String, String)
                                practiceWiseFilename.Add(practice.ToUpper.Trim, currentFileName)
                            End If
                        Else
                            OnHeartbeatError(String.Format("Perfect Pair not found. Filenames: 1. {0}, 2.{1}", runningFile, runningPairFile))
                        End If
                        Exit For
                    End If
                Next
                If Not pairFound Then
                    OnHeartbeatError(String.Format("Pair not found. Filename: {0}. Probale reason: Filename not matching.", runningFile))
                End If
            End If
        Next

        If practiceWiseMaxScoreData IsNot Nothing AndAlso practiceWiseMaxScoreData.Count > 0 Then
            OnHeartbeat("Checking for impact change")
            Dim newForm As New frmImpactChange(_availableScoreUpdates)
            newForm.ShowDialog()

            If _availableScoreUpdates IsNot Nothing AndAlso _availableScoreUpdates.Count > 0 Then
                For Each runningPractice In _availableScoreUpdates.Keys
                    If practiceWiseMaxScoreData.ContainsKey(runningPractice) Then
                        Dim currentFileName As String = practiceWiseFilename(runningPractice)
                        Dim currentMonthScoreData As Object(,) = practiceWiseMaxScoreData(runningPractice)
                        Dim practiceData As PracticeDetails = _availableScoreUpdates(runningPractice.Trim.ToUpper)
                        If practiceData.ProjectedEmployeeData.Keys IsNot Nothing AndAlso practiceData.ProjectedEmployeeData.Keys.Count > 0 Then
                            Dim availableImpactChange As Integer = 0
                            For Each runningEmp In practiceData.ProjectedEmployeeData
                                Dim projectedData As ProjectedDetails = practiceData.ProjectedEmployeeData(runningEmp.Key)
                                If projectedData.Progress Then
                                    availableImpactChange += 1
                                End If
                            Next
                            Dim mendatoryChangeCount As Integer = practiceData.CurrentMonthITPiEmployeeCount - availableImpactChange
                            Dim requiredChangePer As Decimal = practiceData.PreviousMonthImpactPercentage + practiceData.RequiredImpactPercentage
                            Dim requiredChangeCount As Integer = Math.Ceiling(practiceData.CurrentMonthTotalEmployeeCount * requiredChangePer / 100)

                            If requiredChangeCount > mendatoryChangeCount Then
                                Dim extraRequired As Integer = requiredChangeCount - mendatoryChangeCount

                                Dim changeDone As Integer = 0
                                Dim empCount As Integer = 0
                                For Each runningEmp In practiceData.ProjectedEmployeeData.OrderByDescending(Function(x)
                                                                                                                Return x.Value.Foundation.FoundationScore + x.Value.ITPi.ITPiScore
                                                                                                            End Function)
                                    empCount += 1
                                    OnHeartbeat(String.Format("Updating Current Month score for impact {0}/{1}. Practice:{2}", empCount, practiceData.ProjectedEmployeeData.Count, runningPractice))
                                    Dim projectedData As ProjectedDetails = practiceData.ProjectedEmployeeData(runningEmp.Key)
                                    If projectedData.Progress Then
                                        If projectedData.ITPi.ITPiScore > 40 AndAlso projectedData.ITPi.ITPiScore <= 65 Then
                                            Dim rawScoreData As List(Of ScoreData) = Nothing
                                            If practiceData.RawScoreUpdateData.ContainsKey(runningEmp.Key) Then
                                                rawScoreData = practiceData.RawScoreUpdateData(runningEmp.Key)
                                            End If
                                            UpdateCurrentMonthScoreForITPiImpact(projectedData.ITPi, currentMonthScoreData, runningEmp.Key, rawScoreData)
                                            changeDone += 1
                                        End If
                                        If projectedData.Foundation.FoundationScore > 30 AndAlso projectedData.Foundation.FoundationScore <= 40 Then
                                            Dim rawScoreData As List(Of ScoreData) = Nothing
                                            If practiceData.RawScoreUpdateData.ContainsKey(runningEmp.Key) Then
                                                rawScoreData = practiceData.RawScoreUpdateData(runningEmp.Key)
                                            End If
                                            UpdateCurrentMonthScoreForFoundationImpact(projectedData.Foundation, currentMonthScoreData, runningEmp.Key, rawScoreData)
                                            changeDone += 1
                                        End If
                                    End If

                                    If changeDone >= extraRequired Then Exit For
                                Next
                            End If
                        End If

                            Dim outputFileName As String = Path.Combine(_outputDirectoryName, Path.GetFileName(currentFileName))
                        If File.Exists(outputFileName) Then File.Delete(outputFileName)
                        Using xl As New ExcelHelper(outputFileName, ExcelHelper.ExcelOpenStatus.OpenAfreshForWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                            AddHandler xl.Heartbeat, AddressOf OnHeartbeat
                            AddHandler xl.WaitingFor, AddressOf OnWaitingFor
                            OnHeartbeat(String.Format("Writting {0}", outputFileName))
                            Dim range As String = xl.GetNamedRange(1, currentMonthScoreData.GetLength(0) - 1, 1, currentMonthScoreData.GetLength(1) - 1)
                            xl.WriteArrayToExcel(currentMonthScoreData, range)
                            xl.SaveExcel()
                        End Using
                    End If
                Next
            End If
        End If
    End Function

    Public Async Function ProcessEmployeeData() As Task
        Await Task.Delay(1).ConfigureAwait(False)
        If Not Directory.Exists(_outputDirectoryName) Then
            Throw New ApplicationException(String.Format("Directory not found. {0}", _outputDirectoryName))
        End If
        Dim skillList As List(Of String) = Nothing
        For Each runningFile In Directory.GetFiles(_inputDirectoryName)
            If runningFile.ToUpper.Contains("BFSI") AndAlso (skillList Is Nothing OrElse Not skillList.Contains(runningFile, StringComparer.OrdinalIgnoreCase)) Then
                Dim pairFound As Boolean = False
                For Each runningPairFile In Directory.GetFiles(_inputDirectoryName)
                    If runningPairFile.ToUpper <> runningFile.ToUpper AndAlso
                        runningPairFile.Remove(runningPairFile.Count - 13).ToUpper = runningFile.Remove(runningFile.Count - 13).ToUpper Then
                        pairFound = True
                        If skillList Is Nothing Then skillList = New List(Of String)
                        skillList.Add(runningFile)
                        skillList.Add(runningPairFile)
                        Dim pairFileMonth As String = runningPairFile.Substring(runningPairFile.Count - 13, 3)
                        Dim mainFileMonth As String = runningFile.Substring(runningFile.Count - 13, 3)
                        Dim pairFilePreviousMonth As String = Nothing
                        Dim mainFilePreviousMonth As String = Nothing
                        If monthList.ContainsKey(pairFileMonth.ToUpper) Then pairFilePreviousMonth = monthList(pairFileMonth.ToUpper)
                        If monthList.ContainsKey(mainFileMonth.ToUpper) Then mainFilePreviousMonth = monthList(mainFileMonth.ToUpper)

                        Dim currentFileName As String = Nothing
                        Dim previousFileName As String = Nothing
                        If pairFilePreviousMonth IsNot Nothing AndAlso pairFilePreviousMonth.ToUpper = mainFileMonth.ToUpper Then
                            currentFileName = runningPairFile
                            previousFileName = runningFile
                        ElseIf mainFilePreviousMonth IsNot Nothing AndAlso mainFilePreviousMonth.ToUpper = pairFileMonth.ToUpper Then
                            currentFileName = runningFile
                            previousFileName = runningPairFile
                        End If

                        If currentFileName IsNot Nothing AndAlso previousFileName IsNot Nothing Then
                            If currentFileName.ToUpper.Contains("BFSI") AndAlso previousFileName.ToUpper.Contains("BFSI") Then
                                Dim currentMonthEmpData As Object(,) = Nothing
                                Dim previousMonthEmpData As Object(,) = Nothing
                                OnHeartbeat(String.Format("Opening {0}", currentFileName))
                                Using xl As New ExcelHelper(currentFileName, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                                    AddHandler xl.Heartbeat, AddressOf OnHeartbeat
                                    AddHandler xl.WaitingFor, AddressOf OnWaitingFor
                                    Dim allSheets As List(Of String) = xl.GetExcelSheetsName
                                    If allSheets IsNot Nothing AndAlso allSheets.Count > 0 Then
                                        For Each runningSheet In allSheets
                                            If runningSheet.ToUpper.Contains("BFSI") Then
                                                xl.SetActiveSheet(runningSheet)
                                                OnHeartbeat(String.Format("Checking schema {0}", currentFileName))
                                                xl.CheckExcelSchema(empFileSchema.Values.ToArray)
                                                xl.UnFilterSheet(runningSheet)
                                                OnHeartbeat(String.Format("Reading {0}", currentFileName))
                                                currentMonthEmpData = xl.GetExcelInMemory()
                                                Exit For
                                            End If
                                        Next
                                    End If
                                End Using
                                OnHeartbeat(String.Format("Opening {0}", previousFileName))
                                Using xl As New ExcelHelper(previousFileName, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                                    AddHandler xl.Heartbeat, AddressOf OnHeartbeat
                                    AddHandler xl.WaitingFor, AddressOf OnWaitingFor
                                    Dim allSheets As List(Of String) = xl.GetExcelSheetsName
                                    If allSheets IsNot Nothing AndAlso allSheets.Count > 0 Then
                                        For Each runningSheet In allSheets
                                            If runningSheet.ToUpper.Contains("BFSI") Then
                                                xl.SetActiveSheet(runningSheet)
                                                OnHeartbeat(String.Format("Checking schema {0}", previousFileName))
                                                xl.CheckExcelSchema(empFileSchema.Values.ToArray)
                                                If Not allSheets.Contains("Account Validation") Then
                                                    Throw New ApplicationException("'Account Validation' sheet not available in previous month BFSI file")
                                                End If
                                                xl.UnFilterSheet(runningSheet)
                                                OnHeartbeat(String.Format("Reading {0}", previousFileName))
                                                previousMonthEmpData = xl.GetExcelInMemory()
                                                Exit For
                                            End If
                                        Next
                                    End If
                                End Using
                                If currentMonthEmpData IsNot Nothing AndAlso previousMonthEmpData IsNot Nothing Then
                                    Dim currentEmpIdColumnNumber As Integer = _cmn.GetColumnOf2DArray(currentMonthEmpData, 1, empFileSchema("Emp No"))
                                    Dim previousEmpIdColumnNumber As Integer = _cmn.GetColumnOf2DArray(previousMonthEmpData, 1, empFileSchema("Emp No"))
                                    Dim currentLastMonthLevelColumnNumber As Integer = _cmn.GetColumnOf2DArray(currentMonthEmpData, 1, empFileSchema("Last Month Level"))
                                    Dim previousStatusWithFoundationColumnNumber As Integer = _cmn.GetColumnOf2DArray(previousMonthEmpData, 1, empFileSchema("Status With Foundation"))
                                    Dim currentStatusWithFoundationColumnNumber As Integer = _cmn.GetColumnOf2DArray(currentMonthEmpData, 1, empFileSchema("Status With Foundation"))
                                    Dim currentStatusWithoutFoundationColumnNumber As Integer = _cmn.GetColumnOf2DArray(currentMonthEmpData, 1, empFileSchema("Status Without Foundation"))
                                    Dim currentPiApproachColumnNumber As Integer = _cmn.GetColumnOf2DArray(currentMonthEmpData, 1, empFileSchema("Pi Approach"))
                                    Dim currentExperienceOkColumnNumber As Integer = _cmn.GetColumnOf2DArray(currentMonthEmpData, 1, empFileSchema("Experience Ok"))
                                    Dim currentExperienceColumnNumber As Integer = _cmn.GetColumnOf2DArray(currentMonthEmpData, 1, empFileSchema("Experience"))
                                    Dim currentAccountRemarksColumnNumber As Integer = _cmn.GetColumnOf2DArray(currentMonthEmpData, 1, empFileSchema("Account Remarks"))
                                    Dim currentAccountStatusColumnNumber As Integer = _cmn.GetColumnOf2DArray(currentMonthEmpData, 1, empFileSchema("Account Status"))
                                    Dim currentFinalAccountValidationColumnNumber As Integer = _cmn.GetColumnOf2DArray(currentMonthEmpData, 1, empFileSchema("Final Account Validation"))

                                    Dim empIdColumnName As String = _cmn.GetColumnName(currentEmpIdColumnNumber)
                                    Dim experienceColumnName As String = _cmn.GetColumnName(currentExperienceColumnNumber)
                                    Dim accountStatusColumnName As String = _cmn.GetColumnName(currentAccountStatusColumnNumber)
                                    For rowCounter As Integer = 2 To currentMonthEmpData.GetLength(0) - 1
                                        OnHeartbeat(String.Format("Filling Details {0}/{1}. File:{2}", rowCounter, currentMonthEmpData.GetLength(0) - 1, currentFileName))
                                        _cts.Token.ThrowIfCancellationRequested()
                                        Dim empID As String = currentMonthEmpData(rowCounter, currentEmpIdColumnNumber)
                                        If empID IsNot Nothing Then
                                            Dim previousEmpRow As Integer = _cmn.GetRowOf2DArray(previousMonthEmpData, previousEmpIdColumnNumber, empID, True)
                                            If previousEmpRow <> Integer.MinValue Then
                                                If IsNumeric(previousMonthEmpData(previousEmpRow, previousStatusWithFoundationColumnNumber)) Then
                                                    currentMonthEmpData(rowCounter, currentLastMonthLevelColumnNumber) = ""
                                                Else
                                                    currentMonthEmpData(rowCounter, currentLastMonthLevelColumnNumber) = previousMonthEmpData(previousEmpRow, previousStatusWithFoundationColumnNumber)
                                                End If
                                            End If
                                            currentMonthEmpData(rowCounter, currentStatusWithFoundationColumnNumber) = String.Format("=VLOOKUP(${0}{1},'Jscore'!$A:$D,2,FALSE)", empIdColumnName, rowCounter)
                                            currentMonthEmpData(rowCounter, currentStatusWithoutFoundationColumnNumber) = String.Format("=VLOOKUP(${0}{1},'Jscore'!$A:$D,3,FALSE)", empIdColumnName, rowCounter)
                                            currentMonthEmpData(rowCounter, currentPiApproachColumnNumber) = String.Format("=VLOOKUP(${0}{1},'Jscore'!$A:$D,4,FALSE)", empIdColumnName, rowCounter)
                                            currentMonthEmpData(rowCounter, currentExperienceOkColumnNumber) = String.Format("=IF(OR(LEFT({0}{1},2)=""1Y"",LEFT({0}{1},2)=""0Y""),""N"",""Y"")", experienceColumnName, rowCounter)
                                            currentMonthEmpData(rowCounter, currentAccountRemarksColumnNumber) = String.Format("=VLOOKUP(${0}{1},'Account Validation'!$A:$D,3,FALSE)", empIdColumnName, rowCounter)
                                            currentMonthEmpData(rowCounter, currentAccountStatusColumnNumber) = String.Format("=VLOOKUP(${0}{1},'Account Validation'!$A:$D,4,FALSE)", empIdColumnName, rowCounter)
                                            currentMonthEmpData(rowCounter, currentFinalAccountValidationColumnNumber) = String.Format("=IF(ISERROR({0}{1}),""Pending"",IF(OR({0}{1}=""Rotatable"",{0}{1}=""Rotatable > 6 months""),""Eligible"",IF({0}{1}=""Rotated"",""Rotated"",IF({0}{1}=""Pending Validation"",""Pending"",""Not Eligible""))))", accountStatusColumnName, rowCounter)
                                        End If
                                    Next

                                    OnHeartbeat(String.Format("Opening employee file for writting"))
                                    Using xl As New ExcelHelper(currentFileName, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                                        AddHandler xl.Heartbeat, AddressOf OnHeartbeat
                                        AddHandler xl.WaitingFor, AddressOf OnWaitingFor
                                        Dim allSheets As List(Of String) = xl.GetExcelSheetsName
                                        If allSheets IsNot Nothing AndAlso Not allSheets.Contains("JScore") Then xl.CreateNewSheet("JScore")
                                        OnHeartbeat("Trying to copy 'Account Validation' sheet from previous month file")
                                        xl.CopyExcelSheet(previousFileName, "Account Validation")
                                        OnHeartbeat("Trying to copy 'Progress' sheet from previous month file")
                                        xl.CopyExcelSheet(previousFileName, "Progress")
                                        allSheets = xl.GetExcelSheetsName
                                        If allSheets IsNot Nothing AndAlso allSheets.Count > 0 Then
                                            For Each runningSheet In allSheets
                                                If runningSheet.ToUpper.Contains("BFSI") Then
                                                    xl.SetActiveSheet(runningSheet)
                                                    xl.UnFilterSheet(runningSheet)
                                                    OnHeartbeat(String.Format("Writting {0}", currentFileName))
                                                    Dim range As String = xl.GetNamedRange(1, currentMonthEmpData.GetLength(0) - 1, 1, currentMonthEmpData.GetLength(1) - 1)
                                                    xl.WriteArrayToExcel(currentMonthEmpData, range)
                                                    xl.SaveExcel()
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                    End Using
                                    Dim outputFileName As String = Path.Combine(_outputDirectoryName, Path.GetFileName(currentFileName))
                                    If File.Exists(outputFileName) Then File.Delete(outputFileName)
                                    File.Copy(currentFileName, outputFileName)
                                End If
                            End If
                        Else
                            OnHeartbeatError(String.Format("Perfect Pair not found. Filenames: 1. {0}, 2.{1}", runningFile, runningPairFile))
                        End If
                        Exit For
                    End If
                Next
                If Not pairFound Then
                    OnHeartbeatError(String.Format("Pair not found. Filename: {0}. Probale reason: Filename not matching.", runningFile))
                End If
            End If
        Next
    End Function

#Region "Private Functions"
    Private Sub UpdateCurrentMonthScoreForFoundationImpact(ByRef foundationData As FoundationDetails, ByRef currentMonthScoreData As Object(,), ByVal empID As String, ByVal rawScoreData As List(Of ScoreData))
        If foundationData IsNot Nothing AndAlso foundationData.TowerBucketData IsNot Nothing AndAlso foundationData.TowerBucketData.Count > 0 Then
            Dim expectedScore As Decimal = 41
            While foundationData.FoundationScore < expectedScore
                For Each runningTower In foundationData.TowerBucketData.Values.OrderByDescending(Function(x)
                                                                                                     Return x.CurrentMonthModulatedScore
                                                                                                 End Function)
                    Dim wiproSkillColumnList As List(Of String) = runningTower.MappingData.WiproSkillColumnList
                    If wiproSkillColumnList IsNot Nothing AndAlso wiproSkillColumnList.Count > 0 Then
                        Dim skill As String = Nothing
                        If rawScoreData IsNot Nothing AndAlso rawScoreData.Count > 0 Then
                            For Each rawScore In rawScoreData.OrderByDescending(Function(x)
                                                                                    Return (x.CurrentMonthScore - x.PreviousMonthScore)
                                                                                End Function)
                                If wiproSkillColumnList.Contains(rawScore.SkillName, StringComparer.OrdinalIgnoreCase) Then
                                    Dim currentScoreRow As Integer = _cmn.GetRowOf2DArray(currentMonthScoreData, 1, empID, True)
                                    Dim currentScoreColumn As Integer = _cmn.GetColumnOf2DArray(currentMonthScoreData, 1, rawScore.SkillName)
                                    Dim currentScore As String = Nothing
                                    If currentScoreColumn <> Integer.MinValue AndAlso currentScoreRow <> Integer.MinValue Then
                                        currentScore = currentMonthScoreData(currentScoreRow, currentScoreColumn)
                                        If IsNumeric(currentScore) AndAlso Val(currentScore) > 0 AndAlso Val(currentScore) < 100 Then
                                            currentMonthScoreData(currentScoreRow, currentScoreColumn) = currentScore + 1

                                            skill = rawScore.SkillName
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next
                        End If
                        Dim skillScore As Dictionary(Of String, Decimal) = Nothing
                        For Each runningColumn In wiproSkillColumnList
                            _cts.Token.ThrowIfCancellationRequested()
                            Dim currentScoreRow As Integer = _cmn.GetRowOf2DArray(currentMonthScoreData, 1, empID, True)
                            Dim currentScoreColumn As Integer = _cmn.GetColumnOf2DArray(currentMonthScoreData, 1, runningColumn)
                            Dim currentScore As String = Nothing
                            If currentScoreColumn <> Integer.MinValue AndAlso currentScoreRow <> Integer.MinValue Then
                                currentScore = currentMonthScoreData(currentScoreRow, currentScoreColumn)
                                If IsNumeric(currentScore.Trim) Then
                                    If skillScore Is Nothing Then skillScore = New Dictionary(Of String, Decimal)
                                    If Not skillScore.ContainsKey(runningColumn) Then skillScore.Add(runningColumn, currentScore)
                                End If
                            End If
                        Next
                        If skillScore IsNot Nothing AndAlso skillScore.Count > 0 Then
                            For Each runningSkillScore In skillScore.OrderByDescending(Function(x)
                                                                                           Return x.Value
                                                                                       End Function)
                                If skill Is Nothing OrElse runningSkillScore.Key.ToUpper <> skill.ToUpper Then
                                    Dim currentScoreRow As Integer = _cmn.GetRowOf2DArray(currentMonthScoreData, 1, empID, True)
                                    Dim currentScoreColumn As Integer = _cmn.GetColumnOf2DArray(currentMonthScoreData, 1, runningSkillScore.Key)
                                    Dim currentScore As String = Nothing
                                    If currentScoreColumn <> Integer.MinValue AndAlso currentScoreRow <> Integer.MinValue Then
                                        currentScore = currentMonthScoreData(currentScoreRow, currentScoreColumn)
                                        If IsNumeric(currentScore) AndAlso Val(currentScore) > 0 AndAlso Val(currentScore) < 100 Then
                                            currentMonthScoreData(currentScoreRow, currentScoreColumn) = currentScore + 1
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If

                    Dim currentMonthModulatedScore As Decimal = Decimal.MinValue
                    If wiproSkillColumnList IsNot Nothing AndAlso wiproSkillColumnList.Count > 0 Then
                        For Each runningColumn In wiproSkillColumnList
                            _cts.Token.ThrowIfCancellationRequested()
                            Dim currentScoreRow As Integer = _cmn.GetRowOf2DArray(currentMonthScoreData, 1, empID, True)
                            Dim currentScoreColumn As Integer = _cmn.GetColumnOf2DArray(currentMonthScoreData, 1, runningColumn)
                            Dim currentScore As String = Nothing
                            If currentScoreColumn <> Integer.MinValue AndAlso currentScoreRow <> Integer.MinValue Then
                                currentScore = currentMonthScoreData(currentScoreRow, currentScoreColumn)
                                If IsNumeric(currentScore.Trim) Then
                                    currentMonthModulatedScore = Math.Max(currentMonthModulatedScore, Val(currentScore.Trim))
                                End If
                            End If
                        Next
                    End If
                    runningTower.CurrentMonthModulatedScore = currentMonthModulatedScore

                    If foundationData.FoundationScore >= expectedScore Then Exit For
                Next
            End While
        End If
    End Sub

    Private Sub UpdateCurrentMonthScoreForITPiImpact(ByRef itPiData As ITPiDetails, ByRef currentMonthScoreData As Object(,), ByVal empID As String, ByVal rawScoreData As List(Of ScoreData))
        If itPiData IsNot Nothing AndAlso itPiData.TowerBucketData IsNot Nothing AndAlso itPiData.TowerBucketData.Count > 0 Then
            Dim expectedScore As Decimal = 66
            For Each runningTower In itPiData.TowerBucketData.Values.OrderByDescending(Function(x)
                                                                                           Return x.CurrentMonthModulatedScore
                                                                                       End Function)
                Dim wiproSkillColumnList As List(Of String) = runningTower.MappingData.WiproSkillColumnList
                If wiproSkillColumnList IsNot Nothing AndAlso wiproSkillColumnList.Count > 0 Then
                    Dim skill As String = Nothing
                    If rawScoreData IsNot Nothing AndAlso rawScoreData.Count > 0 Then
                        For Each rawScore In rawScoreData.OrderByDescending(Function(x)
                                                                                Return (x.CurrentMonthScore - x.PreviousMonthScore)
                                                                            End Function)
                            If wiproSkillColumnList.Contains(rawScore.SkillName, StringComparer.OrdinalIgnoreCase) Then
                                Dim currentScoreRow As Integer = _cmn.GetRowOf2DArray(currentMonthScoreData, 1, empID, True)
                                Dim currentScoreColumn As Integer = _cmn.GetColumnOf2DArray(currentMonthScoreData, 1, rawScore.SkillName)
                                Dim currentScore As String = Nothing
                                If currentScoreColumn <> Integer.MinValue AndAlso currentScoreRow <> Integer.MinValue Then
                                    currentScore = currentMonthScoreData(currentScoreRow, currentScoreColumn)
                                    If IsNumeric(currentScore) AndAlso Val(currentScore) > 0 Then
                                        currentMonthScoreData(currentScoreRow, currentScoreColumn) = expectedScore

                                        skill = rawScore.SkillName
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                    End If
                    Dim skillScore As Dictionary(Of String, Decimal) = Nothing
                    For Each runningColumn In wiproSkillColumnList
                        _cts.Token.ThrowIfCancellationRequested()
                        Dim currentScoreRow As Integer = _cmn.GetRowOf2DArray(currentMonthScoreData, 1, empID, True)
                        Dim currentScoreColumn As Integer = _cmn.GetColumnOf2DArray(currentMonthScoreData, 1, runningColumn)
                        Dim currentScore As String = Nothing
                        If currentScoreColumn <> Integer.MinValue AndAlso currentScoreRow <> Integer.MinValue Then
                            currentScore = currentMonthScoreData(currentScoreRow, currentScoreColumn)
                            If IsNumeric(currentScore.Trim) Then
                                If skillScore Is Nothing Then skillScore = New Dictionary(Of String, Decimal)
                                If Not skillScore.ContainsKey(runningColumn) Then skillScore.Add(runningColumn, currentScore)
                            End If
                        End If
                    Next
                    If skillScore IsNot Nothing AndAlso skillScore.Count > 0 Then
                        For Each runningSkillScore In skillScore.OrderByDescending(Function(x)
                                                                                       Return x.Value
                                                                                   End Function)
                            If skill Is Nothing OrElse runningSkillScore.Key.ToUpper <> skill.ToUpper Then
                                Dim currentScoreRow As Integer = _cmn.GetRowOf2DArray(currentMonthScoreData, 1, empID, True)
                                Dim currentScoreColumn As Integer = _cmn.GetColumnOf2DArray(currentMonthScoreData, 1, runningSkillScore.Key)
                                Dim currentScore As String = Nothing
                                If currentScoreColumn <> Integer.MinValue AndAlso currentScoreRow <> Integer.MinValue Then
                                    currentScore = currentMonthScoreData(currentScoreRow, currentScoreColumn)
                                    If IsNumeric(currentScore) AndAlso Val(currentScore) > 0 Then
                                        currentMonthScoreData(currentScoreRow, currentScoreColumn) = expectedScore
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
                Exit For
            Next
        End If
    End Sub

    Private Sub UpdateCurrentMonthModulatedScore(ByRef towerData As Dictionary(Of String, TowerDetails), ByVal currentMonthScoreData As Object(,), ByVal empID As String)
        If towerData IsNot Nothing AndAlso towerData.Count > 0 Then
            For Each runningTower In towerData.Values
                Dim currentMonthModulatedScore As Decimal = Decimal.MinValue
                Dim wiproSkillColumnList As List(Of String) = runningTower.MappingData.WiproSkillColumnList
                If wiproSkillColumnList IsNot Nothing AndAlso wiproSkillColumnList.Count > 0 Then
                    For Each runningColumn In wiproSkillColumnList
                        _cts.Token.ThrowIfCancellationRequested()
                        Dim currentScoreRow As Integer = _cmn.GetRowOf2DArray(currentMonthScoreData, 1, empID, True)
                        Dim currentScoreColumn As Integer = _cmn.GetColumnOf2DArray(currentMonthScoreData, 1, runningColumn)
                        Dim currentScore As String = Nothing
                        If currentScoreColumn <> Integer.MinValue AndAlso currentScoreRow <> Integer.MinValue Then
                            currentScore = currentMonthScoreData(currentScoreRow, currentScoreColumn)
                            If IsNumeric(currentScore.Trim) Then
                                currentMonthModulatedScore = Math.Max(currentMonthModulatedScore, Val(currentScore.Trim))
                            End If
                        End If
                    Next
                End If
                runningTower.CurrentMonthModulatedScore = currentMonthModulatedScore
            Next
        End If
    End Sub

    Private Function GetFoundationList(ByVal mappingData As Object(,)) As List(Of String)
        Dim ret As List(Of String) = Nothing
        If mappingData IsNot Nothing Then
            For rowCounter As Integer = 2 To mappingData.GetLength(0) - 1
                For columnCounter As Integer = 6 To mappingData.GetLength(1) - 1
                    If mappingData(rowCounter, columnCounter) IsNot Nothing AndAlso mappingData(rowCounter, columnCounter) <> "" Then
                        If ret Is Nothing Then ret = New List(Of String)
                        ret.Add(mappingData(rowCounter, columnCounter))
                    End If
                Next
                If mappingData(rowCounter, 2) IsNot Nothing AndAlso mappingData(rowCounter, 2) = "I T Pi" Then
                    Exit For
                End If
            Next
        End If
        Return ret
    End Function
#End Region

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
