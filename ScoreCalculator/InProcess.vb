Imports System.IO
Imports Utilities.DAL
Imports System.Threading
Public Class InProcess
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


#Region "Proficiency Levels"
    Private ReadOnly L1 As Decimal = 40.01      'Less Than
    Private ReadOnly L2 As Decimal = 65.01      'Less Than
    Private ReadOnly L3 As Decimal = 85.01      'Less Than
    Private ReadOnly L4 As Decimal = 85         'Greater Than
    Private ReadOnly ISkillQualifyingScore As Decimal = 40.01
#End Region

#Region "Mapping File"
    Private ReadOnly mappingFileSchema As Dictionary(Of String, String)
#End Region

#Region "Score File"
    Private ReadOnly scoreFileSchema As Dictionary(Of String, String)
#End Region

#Region "Employee Details File"
    Private ReadOnly empFileSchema As Dictionary(Of String, String)
#End Region

    Private ReadOnly _cts As CancellationTokenSource
    Private ReadOnly _cmn As Common
    Private ReadOnly _mappingFile As String
    Private ReadOnly _employeeFile As String
    Private ReadOnly _asgFile As String
    Private ReadOnly _directoryName As String
    Private _totalColumn As Integer
    Private _totalRow As Integer
    Private _pivotDataSheet As String
    Private ReadOnly _dataStartingRow As Integer = 11
    Private ReadOnly _outputDirectoryName As String

    Public Sub New(ByVal canceller As CancellationTokenSource, ByVal mappingFile As String, ByVal employeeFile As String, ByVal asgFile As String)
        _cts = canceller
        _mappingFile = mappingFile
        _employeeFile = employeeFile
        _asgFile = asgFile
        _directoryName = Path.GetDirectoryName(_employeeFile)

        _outputDirectoryName = Path.Combine(My.Application.Info.DirectoryPath, "Excel Test", "Post Process")

        _cmn = New Common(_cts)
        mappingFileSchema = New Dictionary(Of String, String) From
            {{"WFT Practice", "WFT Practice"},
            {"WFT Skill Level", "WFT Skill Level"},
            {"WFT Skill Bucket", "WFT Skill Bucket"},
            {"WFT Weightage", "WFT Weightage"},
            {"WFT Subskills", "WFT Subskills"}}

        scoreFileSchema = New Dictionary(Of String, String) From
            {{"Emp No", "Emp No"}, {"Role", "Role"}}

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
            {"Last Month Level", "Last Month Level"}}
    End Sub

    Public Async Function ProcessData() As Task
        Await Task.Delay(1).ConfigureAwait(False)
        Try
            Dim asgData As List(Of ASG) = Nothing
            If File.Exists(_asgFile) Then
                OnHeartbeat("Open ASG file")
                Using xl As New ExcelHelper(_asgFile, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                    AddHandler xl.Heartbeat, AddressOf OnHeartbeat
                    AddHandler xl.WaitingFor, AddressOf OnWaitingFor
                    OnHeartbeat("Reading Employee data")
                    Dim asg As Object(,) = xl.GetExcelInMemory()
                    If asg IsNot Nothing Then
                        For rowCtr As Integer = 2 To asg.GetLength(0) - 1
                            Dim assesmentName As String = asg(rowCtr, 4)
                            Dim asgDetails As ASG = New ASG With {
                            .EmpID = asg(rowCtr, 1),
                            .WFTPractice = assesmentName.Trim.Split("-")(0),
                            .WFTSkillLevel = assesmentName.Trim.Split("-")(1),
                            .WFTSkillBucket = assesmentName.Trim.Split("-")(2),
                            .ProjectedProficiency = assesmentName.Trim.Split("-")(3)
                        }
                            If asgData Is Nothing Then asgData = New List(Of ASG)
                            asgData.Add(asgDetails)
                        Next
                    End If
                End Using
            End If

            Dim employeeData As Object(,) = Nothing
            If File.Exists(_employeeFile) Then
                OnHeartbeat("Open employee details")
                Using xl As New ExcelHelper(_employeeFile, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                    AddHandler xl.Heartbeat, AddressOf OnHeartbeat
                    AddHandler xl.WaitingFor, AddressOf OnWaitingFor
                    Dim allSheets As List(Of String) = xl.GetExcelSheetsName
                    If allSheets IsNot Nothing AndAlso allSheets.Count > 0 Then
                        For Each runningSheet In allSheets
                            If runningSheet.ToUpper.Contains("BFSI") Then
                                xl.SetActiveSheet(runningSheet)
                                OnHeartbeat(String.Format("Checking schema {0}", _employeeFile))
                                xl.CheckExcelSchema(empFileSchema.Values.ToArray)
                                OnHeartbeat("Reading Employee data")
                                employeeData = xl.GetExcelInMemory()
                                Exit For
                            End If
                        Next
                    End If
                End Using

                Dim empOutputFile As String = Path.Combine(_outputDirectoryName, Path.GetFileName(_employeeFile))
                If File.Exists(empOutputFile) Then File.Delete(empOutputFile)
                File.Copy(_employeeFile, empOutputFile)
            Else
                Throw New ApplicationException("Employee file missing")
            End If

            If File.Exists(_mappingFile) Then
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
                                Dim skillRowColumnNumber As KeyValuePair(Of Integer, Integer) = xl.FindAll(mappingFileSchema("WFT Practice"), xl.GetNamedRange(1, 256, 1, 256), True).FirstOrDefault
                                Dim skillName As String = xl.GetData(skillRowColumnNumber.Key + 1, skillRowColumnNumber.Value)
                                Dim scoreFilename As String = Nothing
                                Dim counter As Integer = 0
                                For Each runningFile In Directory.GetFiles(_directoryName, "*.xlsx")
                                    _cts.Token.ThrowIfCancellationRequested()
                                    If runningFile.ToUpper.Contains(skillName.ToUpper) Then
                                        scoreFilename = runningFile
                                        counter += 1
                                        If counter > 1 Then
                                            OnHeartbeatError(String.Format("More than 1 score file exists for skill: {0}", skillName))
                                        End If
                                    End If
                                Next
                                If scoreFilename IsNot Nothing AndAlso counter = 1 Then
                                    Dim skillScoreData As Object(,) = Nothing
                                    Using scoreExcel As New ExcelHelper(scoreFilename, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                                        AddHandler scoreExcel.Heartbeat, AddressOf OnHeartbeat
                                        AddHandler scoreExcel.WaitingFor, AddressOf OnWaitingFor
                                        Try
                                            OnHeartbeat(String.Format("Checking schema {0}", scoreFilename))
                                            scoreExcel.CheckExcelSchema(scoreFileSchema.Values.ToArray)
                                            OnHeartbeat(String.Format("Reading Score data {0}", scoreFilename))
                                            skillScoreData = scoreExcel.GetExcelInMemory()
                                        Catch ex As Exception
                                            OnHeartbeatError(ex.Message)
                                        End Try
                                    End Using

                                    If skillScoreData IsNot Nothing Then
                                        Dim skillOutputFileName As String = Path.Combine(_directoryName, String.Format("{0} Output File {1}.xlsx", skillName, scoreFilename.Substring(scoreFilename.Count - 13, 8).ToUpper))
                                        If File.Exists(skillOutputFileName) Then File.Delete(skillOutputFileName)

                                        Dim skillMaxScore As Dictionary(Of String, Object(,)) = Nothing
                                        Dim bestScoreData As Dictionary(Of String, List(Of String)) = Nothing
                                        Using outputExcel As New ExcelHelper(skillOutputFileName, ExcelHelper.ExcelOpenStatus.OpenAfreshForWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                                            'AddHandler outputExcel.Heartbeat, AddressOf OnHeartbeat
                                            'AddHandler outputExcel.WaitingFor, AddressOf OnWaitingFor
                                            Dim subSkillColumnNumber As Integer = xl.FindAll(mappingFileSchema("WFT Subskills"), xl.GetNamedRange(1, 256, 1, 256), True).FirstOrDefault.Value
                                            Dim subSkillLastRow As Integer = xl.GetLastRow(subSkillColumnNumber)
                                            _totalColumn = subSkillLastRow
                                            _totalRow = skillScoreData.GetLength(0)
                                            For subskillRow As Integer = 2 To subSkillLastRow
                                                _cts.Token.ThrowIfCancellationRequested()
                                                Dim subskill As String = xl.GetData(subskillRow, subSkillColumnNumber)
                                                If subskill IsNot Nothing Then
                                                    OnHeartbeat(String.Format("Creating max score sheet for {0} in {1}", subskill, skillName))
                                                    Dim subSkillData(skillScoreData.GetLength(0), 0) As Object
                                                    Dim rowCtr As Integer = 0
                                                    Dim colCtr As Integer = 0

                                                    If colCtr > UBound(subSkillData, 2) Then ReDim Preserve subSkillData(UBound(subSkillData, 1), 0 To UBound(subSkillData, 2) + 1)
                                                    subSkillData(rowCtr, colCtr) = "Emp No"
                                                    colCtr += 1
                                                    If colCtr > UBound(subSkillData, 2) Then ReDim Preserve subSkillData(UBound(subSkillData, 1), 0 To UBound(subSkillData, 2) + 1)
                                                    subSkillData(rowCtr, colCtr) = "Max Score"
                                                    colCtr += 1

                                                    Dim lastColumn As Integer = xl.GetLastCol(subskillRow)
                                                    For wiproSkillColumn As Integer = subSkillColumnNumber + 1 To lastColumn
                                                        _cts.Token.ThrowIfCancellationRequested()
                                                        rowCtr = 0
                                                        Dim wiproSkill As String = xl.GetData(subskillRow, wiproSkillColumn)
                                                        If wiproSkill IsNot Nothing Then
                                                            If colCtr > UBound(subSkillData, 2) Then ReDim Preserve subSkillData(UBound(subSkillData, 1), 0 To UBound(subSkillData, 2) + 1)
                                                            subSkillData(rowCtr, colCtr) = wiproSkill
                                                            Dim skillScoreColumnNumber As Integer = _cmn.GetColumnOf2DArray(skillScoreData, 1, wiproSkill)
                                                            Dim empColumnNumber As Integer = _cmn.GetColumnOf2DArray(skillScoreData, 1, scoreFileSchema("Emp No"))
                                                            For empRow As Integer = 2 To skillScoreData.GetLength(0)
                                                                _cts.Token.ThrowIfCancellationRequested()
                                                                rowCtr += 1
                                                                Dim empScore As Integer = 0
                                                                If skillScoreColumnNumber <> Integer.MinValue AndAlso
                                                                    skillScoreData(empRow, skillScoreColumnNumber) IsNot Nothing Then
                                                                    empScore = skillScoreData(empRow, skillScoreColumnNumber)
                                                                End If
                                                                subSkillData(rowCtr, 0) = skillScoreData(empRow, empColumnNumber)
                                                                subSkillData(rowCtr, 1) = String.Format("=MAX(C{0}:{1}{2})", rowCtr + 1, outputExcel.GetColumnName(1000), rowCtr + 1)
                                                                subSkillData(rowCtr, colCtr) = empScore
                                                            Next
                                                            colCtr += 1
                                                        End If
                                                    Next

                                                    _cts.Token.ThrowIfCancellationRequested()
                                                    Dim loadedSheets As List(Of String) = outputExcel.GetExcelSheetsName()
                                                    Dim sheetName As String = GetSheetName(loadedSheets, subskill)
                                                    outputExcel.CreateNewSheet(sheetName)
                                                    If outputExcel.SetActiveSheet(sheetName) Then
                                                        If skillMaxScore Is Nothing Then skillMaxScore = New Dictionary(Of String, Object(,))
                                                        skillMaxScore.Add(sheetName, subSkillData)
                                                        Dim range As String = outputExcel.GetNamedRange(1, subSkillData.GetLength(0) - 1, 1, subSkillData.GetLength(1) - 1)
                                                        _cts.Token.ThrowIfCancellationRequested()
                                                        outputExcel.WriteArrayToExcel(subSkillData, range)
                                                        _cts.Token.ThrowIfCancellationRequested()
                                                        Erase subSkillData
                                                        subSkillData = Nothing
                                                        outputExcel.SaveExcel()
                                                    End If
                                                End If
                                            Next

                                            Dim wftSubSkillList As List(Of String) = Nothing
                                            'OnHeartbeat("Creating summary sheet")
                                            Dim skillLevelColumnNumber As Integer = xl.FindAll(mappingFileSchema("WFT Skill Level"), xl.GetNamedRange(1, 256, 1, 256), True).FirstOrDefault.Value
                                            Dim skillLevelLastRow As Integer = xl.GetLastRow(skillLevelColumnNumber)
                                            Dim lastSkillLevel As String = Nothing
                                            Dim previousSkillLevel As String = Nothing
                                            Dim foundationTotalProficiencyColumnNumber As Integer = Integer.MinValue
                                            For skillLevelRow As Integer = 2 To skillLevelLastRow
                                                _cts.Token.ThrowIfCancellationRequested()
                                                Dim skillLevel As String = xl.GetData(skillLevelRow, skillLevelColumnNumber)
                                                If skillLevel IsNot Nothing Then
                                                    If lastSkillLevel Is Nothing OrElse lastSkillLevel.ToUpper <> skillLevel.ToUpper Then
                                                        previousSkillLevel = lastSkillLevel
                                                        lastSkillLevel = skillLevel
                                                        outputExcel.CreateNewSheet(skillLevel)
                                                        outputExcel.SetActiveSheet(skillLevel)

                                                        Dim empDataColumns As Dictionary(Of String, Integer) = Nothing
                                                        If empFileSchema IsNot Nothing AndAlso empFileSchema.Count > 0 Then
                                                            outputExcel.SetData(1, 1, "Proficiency Level Assumptions")
                                                            outputExcel.SetData(2, 1, "L1")
                                                            outputExcel.SetData(2, 2, "Less Than (<)")
                                                            outputExcel.SetData(2, 3, L1, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                                            outputExcel.SetData(3, 1, "L2")
                                                            outputExcel.SetData(3, 2, "Less Than (<)")
                                                            outputExcel.SetData(3, 3, L2, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                                            outputExcel.SetData(4, 1, "L3")
                                                            outputExcel.SetData(4, 2, "Less Than (<)")
                                                            outputExcel.SetData(4, 3, L3, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                                            outputExcel.SetData(5, 1, "L4")
                                                            outputExcel.SetData(5, 2, "Greater Than (>)")
                                                            outputExcel.SetData(5, 3, L4, "##,##,##0.00", ExcelHelper.XLAlign.Right)

                                                            Dim rowCtr As Integer = 0
                                                            Dim columnCtr As Integer = 0
                                                            OnHeartbeat("Writting headers and mapping employee details columns")
                                                            If skillMaxScore IsNot Nothing AndAlso skillMaxScore.Count > 0 Then
                                                                Dim skillMaxScoreData As Object(,) = skillMaxScore.FirstOrDefault.Value
                                                                If skillMaxScoreData IsNot Nothing Then
                                                                    Dim empDetails(skillMaxScoreData.GetLength(0) - 1, empFileSchema.Count) As Object
                                                                    For Each empData In empFileSchema.Keys
                                                                        _cts.Token.ThrowIfCancellationRequested()
                                                                        If empData.ToUpper <> "Last Month Level".ToUpper Then
                                                                            empDetails(rowCtr, columnCtr) = empData
                                                                        End If
                                                                        Dim empDataColumnName As String = empFileSchema(empData)
                                                                        Dim empDataColumnNumber As Integer = _cmn.GetColumnOf2DArray(employeeData, 1, empDataColumnName)
                                                                        If empDataColumns Is Nothing Then empDataColumns = New Dictionary(Of String, Integer)
                                                                        empDataColumns.Add(empData, empDataColumnNumber)
                                                                        columnCtr += 1
                                                                    Next

                                                                    rowCtr += 1
                                                                    For empRow As Integer = 1 To skillMaxScoreData.GetLength(0) - 1
                                                                        _cts.Token.ThrowIfCancellationRequested()
                                                                        OnHeartbeat(String.Format("Filling employee data. ({0}/{1}), Level:{2}, Skill Name:{3}", empRow, skillMaxScoreData.GetLength(0) - 1, skillLevel, skillName))
                                                                        columnCtr = 0
                                                                        Dim empNo As String = skillMaxScoreData(empRow, 0)
                                                                        If empNo IsNot Nothing Then
                                                                            empDetails(rowCtr, columnCtr) = empNo
                                                                            Dim empDataRowNumber As Integer = _cmn.GetRowOf2DArray(employeeData, empDataColumns("Emp No"), empNo, True)
                                                                            If empDataRowNumber <> Integer.MinValue Then
                                                                                For Each empData In empFileSchema.Keys
                                                                                    If empData.ToUpper <> "Emp No".ToUpper AndAlso
                                                                                        empData.ToUpper <> "Last Month Level".ToUpper Then
                                                                                        columnCtr += 1
                                                                                        Dim empDataColumnNumber As Integer = empDataColumns(empData)
                                                                                        If empDataColumnNumber <> Integer.MinValue Then
                                                                                            empDetails(rowCtr, columnCtr) = employeeData(empDataRowNumber, empDataColumnNumber)
                                                                                        End If
                                                                                    End If
                                                                                Next
                                                                            End If
                                                                            rowCtr += 1
                                                                        End If
                                                                    Next

                                                                    Dim range As String = outputExcel.GetNamedRange(_dataStartingRow, empDetails.GetLength(0) - 1, 1, empDetails.GetLength(1) - 1)
                                                                    outputExcel.WriteArrayToExcel(empDetails, range)
                                                                End If
                                                            End If
                                                        End If
                                                        outputExcel.SaveExcel()

                                                        Dim summaryData(_totalRow, _totalColumn) As Object
                                                        Dim rowNumber As Integer = 1
                                                        Dim columnNumber As Integer = 1
                                                        If empFileSchema IsNot Nothing Then columnNumber = columnNumber + empFileSchema.Count - 1
                                                        Dim skillBucketColumnNumber As Integer = xl.FindAll(mappingFileSchema("WFT Skill Bucket"), xl.GetNamedRange(1, 256, 1, 256), True).FirstOrDefault.Value
                                                        Dim skillBucketLastRow As Integer = xl.GetLastRow(skillBucketColumnNumber)
                                                        For skillBucketRow As Integer = 2 To skillBucketLastRow
                                                            _cts.Token.ThrowIfCancellationRequested()
                                                            Dim skillBucket As String = xl.GetData(skillBucketRow, skillBucketColumnNumber)
                                                            If skillBucket IsNot Nothing Then
                                                                Dim relatedSkillLevel As String = xl.GetData(skillBucketRow, skillLevelColumnNumber)
                                                                If relatedSkillLevel.ToUpper = lastSkillLevel.ToUpper Then
                                                                    Dim nextSkillBucketRow As Integer = skillBucketRow
                                                                    For tempSkillBucketRow As Integer = skillBucketRow + 1 To skillBucketLastRow
                                                                        _cts.Token.ThrowIfCancellationRequested()
                                                                        Dim tempSkillBucket As String = xl.GetData(tempSkillBucketRow, skillBucketColumnNumber)
                                                                        If tempSkillBucket IsNot Nothing Then
                                                                            nextSkillBucketRow = tempSkillBucketRow
                                                                            Exit For
                                                                        End If
                                                                    Next
                                                                    Dim skillWtgColumnNumber As Integer = xl.FindAll(mappingFileSchema("WFT Weightage"), xl.GetNamedRange(1, 256, 1, 256), True).FirstOrDefault.Value
                                                                    Dim skillWtg As String = xl.GetData(skillBucketRow, skillWtgColumnNumber)
                                                                    SetDataTo2DArray(9, columnNumber, String.Format("{0}%", skillWtg * 100), summaryData)
                                                                    SetDataTo2DArray(10, columnNumber, skillBucket, summaryData)
                                                                    Dim groupData As List(Of String) = New List(Of String) From {
                                                                        String.Format("${0}${1}", outputExcel.GetColumnName(columnNumber), 9)
                                                                    }
                                                                    Dim wftSubSkillColumnNumber As Integer = xl.FindAll(mappingFileSchema("WFT Subskills"), xl.GetNamedRange(1, 256, 1, 256), True).FirstOrDefault.Value
                                                                    Dim wftSubSkillLastRow As Integer = xl.GetLastRow(wftSubSkillColumnNumber)
                                                                    If skillBucketRow <> nextSkillBucketRow Then wftSubSkillLastRow = nextSkillBucketRow - 1
                                                                    Dim startColumnNumber As Integer = columnNumber
                                                                    Dim lastRowNumber As Integer = rowNumber
                                                                    For wftSubSkillRow As Integer = skillBucketRow To wftSubSkillLastRow
                                                                        _cts.Token.ThrowIfCancellationRequested()
                                                                        rowNumber = _dataStartingRow
                                                                        Dim wftSubSkill As String = xl.GetData(wftSubSkillRow, wftSubSkillColumnNumber)
                                                                        If wftSubSkill IsNot Nothing Then
                                                                            SetDataTo2DArray(rowNumber, columnNumber, wftSubSkill, summaryData)
                                                                            Dim maxScoreSheetName As String = GetSheetName(wftSubSkillList, wftSubSkill)
                                                                            If wftSubSkillList Is Nothing Then wftSubSkillList = New List(Of String)
                                                                            wftSubSkillList.Add(maxScoreSheetName)
                                                                            Dim maxScoreData As Object(,) = skillMaxScore(maxScoreSheetName)
                                                                            For empNoRow As Integer = 1 To maxScoreData.GetLength(0) - 1
                                                                                _cts.Token.ThrowIfCancellationRequested()
                                                                                OnHeartbeat(String.Format("Filling subskill score. ({0}/{1}). Subskill:{2}({3}/{4}), Level:{5}, Skill Name:{6}",
                                                                                                          empNoRow, maxScoreData.GetLength(0) - 1, wftSubSkill, wftSubSkillRow - skillBucketRow + 1,
                                                                                                          wftSubSkillLastRow - skillBucketRow + 1, skillLevel, skillName))
                                                                                Dim empID As String = maxScoreData(empNoRow, 0)
                                                                                If empID IsNot Nothing Then
                                                                                    rowNumber += 1
                                                                                    SetDataTo2DArray(rowNumber, columnNumber, String.Format("='{0}'!B{1}", maxScoreSheetName, empNoRow + 1), summaryData)

                                                                                    If asgData IsNot Nothing AndAlso asgData.Count > 0 Then
                                                                                        Dim asgDetails As List(Of ASG) = asgData.FindAll(Function(x)
                                                                                                                                             Return x.EmpID = empID
                                                                                                                                         End Function)
                                                                                        If asgData IsNot Nothing AndAlso asgDetails.Count > 0 Then
                                                                                            For Each runningASGData In asgDetails
                                                                                                If Path.GetFileNameWithoutExtension(skillOutputFileName).Contains(runningASGData.WFTPractice) Then
                                                                                                    If runningASGData.WFTSkillLevel = "Foundation" Then
                                                                                                        If skillLevel = "Foundation" Then
                                                                                                            outputExcel.SetActiveSheet(maxScoreSheetName)
                                                                                                            Dim lastMaxScrClm As Integer = outputExcel.GetLastCol(1) + 1
                                                                                                            Dim range As String = outputExcel.GetNamedRange(1, outputExcel.GetLastRow(1), 1, 1)
                                                                                                            Dim rowEmp As Integer = outputExcel.FindAll(empID, range, True).FirstOrDefault.Key
                                                                                                            outputExcel.SetData(rowEmp, lastMaxScrClm, runningASGData.ProjectedScore, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                                                                                        End If
                                                                                                    Else
                                                                                                        If skillBucket.ToUpper = runningASGData.WFTSkillBucket.ToUpper Then
                                                                                                            outputExcel.SetActiveSheet(maxScoreSheetName)
                                                                                                            Dim lastMaxScrClm As Integer = outputExcel.GetLastCol(1) + 1
                                                                                                            Dim range As String = outputExcel.GetNamedRange(1, outputExcel.GetLastRow(1), 1, 1)
                                                                                                            Dim rowEmp As Integer = outputExcel.FindAll(empID, range, True).FirstOrDefault.Key
                                                                                                            outputExcel.SetData(rowEmp, lastMaxScrClm, runningASGData.ProjectedScore, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                                                                                        End If
                                                                                                    End If
                                                                                                End If
                                                                                            Next
                                                                                            outputExcel.SetActiveSheet(skillLevel)
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            Next
                                                                            columnNumber += 1
                                                                            lastRowNumber = rowNumber
                                                                        End If
                                                                    Next
                                                                    For rowCout As Integer = _dataStartingRow To lastRowNumber
                                                                        _cts.Token.ThrowIfCancellationRequested()
                                                                        OnHeartbeat(String.Format("Calculating best score and proficiency. ({0}/{1}). Skill Bucket:{2}, Level:{3}, Skill Name:{4}",
                                                                                                  rowCout - 10, lastRowNumber - 10, skillBucket, skillLevel, skillName))
                                                                        If rowCout = _dataStartingRow Then
                                                                            SetDataTo2DArray(rowCout, columnNumber, "BEST SCORE FROM Group", summaryData)
                                                                            SetDataTo2DArray(rowCout, columnNumber + 1, "Group Proficiency", summaryData)
                                                                        Else
                                                                            SetDataTo2DArray(rowCout, columnNumber, String.Format("=MAX({0}{1}:{2}{3})", outputExcel.GetColumnName(startColumnNumber), rowCout, outputExcel.GetColumnName(columnNumber - 1), rowCout), summaryData)
                                                                            Dim cellNumber As String = String.Format("{0}{1}", outputExcel.GetColumnName(columnNumber), rowCout)
                                                                            SetDataTo2DArray(rowCout, columnNumber + 1, String.Format("=IF({0}<$C$2,""L1"",IF({0}<$C$3,""L2"",IF({0}<$C$4,""L3"",""L4"")))", cellNumber), summaryData)
                                                                        End If
                                                                    Next
                                                                    groupData.Add(outputExcel.GetColumnName(columnNumber))
                                                                    If bestScoreData Is Nothing Then bestScoreData = New Dictionary(Of String, List(Of String))
                                                                    bestScoreData.Add(skillBucket, groupData)
                                                                    columnNumber += 2
                                                                End If
                                                            End If
                                                        Next
                                                        Dim summaryRange As String = outputExcel.GetNamedRange(9, summaryData.GetLength(0) - 1, 1 + (empFileSchema.Count - 1), summaryData.GetLength(1) - 1)
                                                        Try
                                                            outputExcel.WriteArrayToExcel(summaryData, summaryRange)
                                                            outputExcel.SaveExcel()
                                                        Catch ex As Exception
                                                            Throw ex
                                                        End Try

                                                        Dim summaryRowNumber As Integer = _dataStartingRow
                                                        Dim lastSummaryRow As Integer = outputExcel.GetLastRow()
                                                        Dim lastSummaryColumn As Integer = outputExcel.GetLastCol(summaryRowNumber)
                                                        If skillLevel.Contains("Foundation") Then
                                                            Dim startSummaryColumn As Integer = 1 + lastSummaryColumn
                                                            outputExcel.SetData(summaryRowNumber, startSummaryColumn, "Total Proficiency Score")
                                                            outputExcel.SetData(summaryRowNumber, startSummaryColumn + 1, "Total Proficiency Level")
                                                            If bestScoreData IsNot Nothing AndAlso bestScoreData.Count > 0 Then
                                                                For runningSummaryRow As Integer = summaryRowNumber + 1 To lastSummaryRow
                                                                    _cts.Token.ThrowIfCancellationRequested()
                                                                    OnHeartbeat(String.Format("Calculating final scores. ({0}/{1}). Skill Level:{2}, Skill Name:{3}", runningSummaryRow - _dataStartingRow, lastSummaryRow - _dataStartingRow, skillLevel, skillName))
                                                                    Dim totalProficiency As String = ""
                                                                    For Each group In bestScoreData.Values
                                                                        totalProficiency = String.Format("{0}+{1}*{2}{3}", totalProficiency, group(0), group(1), runningSummaryRow)
                                                                    Next
                                                                    totalProficiency = totalProficiency.Substring(1)
                                                                    totalProficiency = String.Format("=SUM({0})", totalProficiency)
                                                                    outputExcel.SetCellFormula(runningSummaryRow, startSummaryColumn, totalProficiency)
                                                                    outputExcel.SetCellFormula(runningSummaryRow, startSummaryColumn + 1, String.Format("=IF({0}<$C$2,""L1"",IF({0}<$C$3,""L2"",IF({0}<$C$4,""L3"",""L4"")))", String.Format("{0}{1}", outputExcel.GetColumnName(startSummaryColumn), runningSummaryRow)))
                                                                Next
                                                            End If
                                                            foundationTotalProficiencyColumnNumber = startSummaryColumn
                                                            outputExcel.SaveExcel()
                                                        ElseIf skillLevel.Contains("I T Pi") Then
                                                            _pivotDataSheet = skillLevel
                                                            If foundationTotalProficiencyColumnNumber <> Integer.MinValue Then
                                                                Dim startSummaryColumn As Integer = 1 + lastSummaryColumn
                                                                outputExcel.SetData(summaryRowNumber - 1, startSummaryColumn, "Foundation Level Proficiency")
                                                                outputExcel.SetData(summaryRowNumber, startSummaryColumn, "Total Proficiency Score")
                                                                outputExcel.SetData(summaryRowNumber, startSummaryColumn + 1, "Total Proficiency Level")

                                                                outputExcel.SetData(summaryRowNumber - 2, startSummaryColumn + 2, "I Skill Qualifying Score")
                                                                outputExcel.SetData(summaryRowNumber - 1, startSummaryColumn + 2, ISkillQualifyingScore, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                                                Dim ISkillQualifyingScoreCell As String = String.Format("${0}${1}", outputExcel.GetColumnName(startSummaryColumn + 2), summaryRowNumber - 1)
                                                                outputExcel.SetData(summaryRowNumber, startSummaryColumn + 2, "'>=L3 proficiency")
                                                                outputExcel.SetData(summaryRowNumber, startSummaryColumn + 3, "'>=L2 proficiency")

                                                                outputExcel.SetData(summaryRowNumber - 2, startSummaryColumn + 4, "I Skill Qualifying Score")
                                                                outputExcel.SetData(summaryRowNumber - 1, startSummaryColumn + 4, 0, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                                                Dim dummyISkillQualifyingScoreCell As String = String.Format("${0}${1}", outputExcel.GetColumnName(startSummaryColumn + 4), summaryRowNumber - 1)
                                                                outputExcel.SetData(summaryRowNumber, startSummaryColumn + 4, "'>=L3 proficiency")
                                                                outputExcel.SetData(summaryRowNumber, startSummaryColumn + 5, "'>=L2 proficiency")

                                                                outputExcel.SetData(summaryRowNumber - 1, startSummaryColumn + 6, "With Foundation")
                                                                outputExcel.SetData(summaryRowNumber, startSummaryColumn + 6, "I/T/Pi")

                                                                outputExcel.SetData(summaryRowNumber - 1, startSummaryColumn + 7, "Without Foundation")
                                                                outputExcel.SetData(summaryRowNumber, startSummaryColumn + 7, "I/T/Pi")
                                                                outputExcel.SetData(summaryRowNumber, startSummaryColumn + 8, "Pi Approach")

                                                                outputExcel.SetData(summaryRowNumber, startSummaryColumn + 9, "Last Month Level")

                                                                Dim firstBestScoreColumn As Integer = outputExcel.FindAll("BEST SCORE FROM Group", outputExcel.GetNamedRange(1, 256, 1, 256)).FirstOrDefault.Value
                                                                For runningSummaryRow As Integer = summaryRowNumber + 1 To lastSummaryRow
                                                                    _cts.Token.ThrowIfCancellationRequested()
                                                                    OnHeartbeat(String.Format("Calculating final scores. ({0}/{1}). Skill Level:{2}, Skill Name:{3}", runningSummaryRow - _dataStartingRow, lastSummaryRow - _dataStartingRow, skillLevel, skillName))
                                                                    outputExcel.SetCellFormula(runningSummaryRow, startSummaryColumn, String.Format("='{0}'!{1}{2}", previousSkillLevel, outputExcel.GetColumnName(foundationTotalProficiencyColumnNumber), runningSummaryRow))
                                                                    Dim foundationTotalScoreCell As String = String.Format("{0}{1}", outputExcel.GetColumnName(startSummaryColumn), runningSummaryRow)
                                                                    outputExcel.SetCellFormula(runningSummaryRow, startSummaryColumn + 1, String.Format("='{0}'!{1}{2}", previousSkillLevel, outputExcel.GetColumnName(foundationTotalProficiencyColumnNumber + 1), runningSummaryRow))

                                                                    Dim calculateCell As String = String.Format("{0}{1}:{2}{3}", outputExcel.GetColumnName(firstBestScoreColumn), runningSummaryRow, outputExcel.GetColumnName(lastSummaryColumn), runningSummaryRow)
                                                                    Dim iProficiencyScore As String = String.Format("=MIN(IF({0}>={1},COUNTIF({2},""L3"")+COUNTIF({3},""L4""),0),2)", foundationTotalScoreCell, ISkillQualifyingScoreCell, calculateCell, calculateCell)
                                                                    outputExcel.SetCellFormula(runningSummaryRow, startSummaryColumn + 2, iProficiencyScore)
                                                                    Dim l3Cell As String = String.Format("{0}{1}", outputExcel.GetColumnName(startSummaryColumn + 2), runningSummaryRow)
                                                                    Dim iProficiencyScore2 As String = String.Format("=IF({0}>={1},COUNTIF({2},""L2"")+COUNTIF({3},""L3"")+COUNTIF({4},""L4""),0)", foundationTotalScoreCell, ISkillQualifyingScoreCell, calculateCell, calculateCell, calculateCell)
                                                                    outputExcel.SetCellFormula(runningSummaryRow, startSummaryColumn + 3, iProficiencyScore2)
                                                                    Dim l2Cell As String = String.Format("{0}{1}", outputExcel.GetColumnName(startSummaryColumn + 3), runningSummaryRow)

                                                                    Dim dummyiProficiencyScore As String = String.Format("=MIN(IF({0}>={1},COUNTIF({2},""L3"")+COUNTIF({3},""L4""),0),2)", foundationTotalScoreCell, dummyISkillQualifyingScoreCell, calculateCell, calculateCell)
                                                                    outputExcel.SetCellFormula(runningSummaryRow, startSummaryColumn + 4, dummyiProficiencyScore)
                                                                    Dim dummyl3Cell As String = String.Format("{0}{1}", outputExcel.GetColumnName(startSummaryColumn + 4), runningSummaryRow)
                                                                    Dim dummyiProficiencyScore2 As String = String.Format("=IF({0}>={1},COUNTIF({2},""L2"")+COUNTIF({3},""L3"")+COUNTIF({4},""L4""),0)", foundationTotalScoreCell, dummyISkillQualifyingScoreCell, calculateCell, calculateCell, calculateCell)
                                                                    outputExcel.SetCellFormula(runningSummaryRow, startSummaryColumn + 5, dummyiProficiencyScore2)
                                                                    Dim dummyl2Cell As String = String.Format("{0}{1}", outputExcel.GetColumnName(startSummaryColumn + 5), runningSummaryRow)

                                                                    Dim withFoundationITPi As String = String.Format("=IF(AND({0}>=2,{1}>=4),""Pi"",IF(AND({0}>=1,{1}>=3),""T"",IF({0}>=1,""I"",IF({2}>={3},""Foundation Complete"",""Foundation Pending""))))", l3Cell, l2Cell, foundationTotalScoreCell, ISkillQualifyingScoreCell)
                                                                    outputExcel.SetCellFormula(runningSummaryRow, startSummaryColumn + 6, withFoundationITPi)
                                                                    Dim withFoundationITPiCell As String = String.Format("{0}{1}", outputExcel.GetColumnName(startSummaryColumn + 6), runningSummaryRow)

                                                                    Dim withoutFoundationITPi As String = String.Format("=IF(AND({0}>=2,{1}>=4),""Pi"",IF(AND({0}>=1,{1}>=3),""T"",IF({0}>=1,""I"",IF({2}>={3},""Foundation Complete"",""Foundation Pending""))))", dummyl3Cell, dummyl2Cell, foundationTotalScoreCell, dummyISkillQualifyingScoreCell)
                                                                    outputExcel.SetCellFormula(runningSummaryRow, startSummaryColumn + 7, withoutFoundationITPi)

                                                                    Dim piApproch As String = String.Format("=IF({0}=""Pi"",""Pi"",IF({0}=""T"",""Near Pi-T"",IF(AND({0}=""I"",{1}>=2),""Near Pi-2I"","""")))", withFoundationITPiCell, l3Cell)
                                                                    outputExcel.SetCellFormula(runningSummaryRow, startSummaryColumn + 8, piApproch)

                                                                    Dim empId As String = outputExcel.GetData(runningSummaryRow, 1)
                                                                    If empId IsNot Nothing AndAlso empId <> "" Then
                                                                        Dim empDetailsRowNumber As Integer = _cmn.GetRowOf2DArray(employeeData, empDataColumns("Emp No"), empId, True)
                                                                        If empDetailsRowNumber <> Integer.MinValue Then
                                                                            outputExcel.SetCellFormula(runningSummaryRow, startSummaryColumn + 9, employeeData(empDetailsRowNumber, empDataColumns("Last Month Level")))
                                                                        End If
                                                                    End If
                                                                Next
                                                            End If
                                                            outputExcel.SaveExcel()
                                                        End If
                                                    End If
                                                End If
                                            Next

                                            OnHeartbeat(String.Format("Creating Pivot for {0}", skillOutputFileName))
                                            Dim dataSheetRange As String = String.Format("{0}{1}:{2}{3}",
                                                                                         outputExcel.GetColumnName(1), _dataStartingRow,
                                                                                         outputExcel.GetColumnName(outputExcel.GetLastCol(_dataStartingRow)), outputExcel.GetLastRow)
                                            outputExcel.CreateNewSheet("Progress")
                                            outputExcel.CreatPivotTable(_pivotDataSheet, dataSheetRange, "Progress", "A3",
                                                                        New List(Of String) From {empFileSchema("Last Month Level")},
                                                                        New List(Of String) From {"I/T/Pi"},
                                                                        New Dictionary(Of String, ExcelHelper.XLFunction) From {{"Emp No", ExcelHelper.XLFunction.Count}},
                                                                        Nothing)

                                            Dim orderList As List(Of String) = New List(Of String) From {"Foundation Pending", "Foundation Complete", "I", "T", "Pi"}
                                            'outputExcel.ReorderPivotTable("Progress", "Last Month Level", orderList)
                                            'outputExcel.ReorderPivotTable("Progress", "I/T/Pi", orderList)

                                            'outputExcel.SetData(12, 1, skillName)
                                            'outputExcel.SetData(12, 2, "# of resources whose levels changed")
                                            'outputExcel.SetData(12, 4, "% of resources whose levels changed")
                                            'outputExcel.SetData(13, 2, "Up")
                                            'outputExcel.SetData(13, 3, "Down")
                                            'outputExcel.SetData(13, 4, "Up")
                                            'outputExcel.SetData(13, 5, "Down")

                                            'outputExcel.SetData(14, 1, "Foundation Complete")
                                            'outputExcel.SetCellFormula(14, 2, "=B6+G6")
                                            'outputExcel.SetCellFormula(14, 3, "=C5")
                                            'outputExcel.SetCellFormula(14, 4, "=B14/$B$19")
                                            'outputExcel.SetCellFormula(14, 5, "=C14/$B$19")

                                            'outputExcel.SetData(15, 1, "I")
                                            'outputExcel.SetCellFormula(15, 2, "=SUM(B7:C7)+G7")
                                            'outputExcel.SetCellFormula(15, 3, "=SUM(D5:D6)")
                                            'outputExcel.SetCellFormula(15, 4, "=B15/$B$19")
                                            'outputExcel.SetCellFormula(15, 5, "=C15/$B$19")

                                            'outputExcel.SetData(16, 1, "T")
                                            'outputExcel.SetCellFormula(16, 2, "=SUM(B8:D8)+G8")
                                            'outputExcel.SetCellFormula(16, 3, "=SUM(E5:E7)")
                                            'outputExcel.SetCellFormula(16, 4, "=B16/$B$19")
                                            'outputExcel.SetCellFormula(16, 5, "=C16/$B$19")

                                            'outputExcel.SetData(17, 1, "Pi")
                                            'outputExcel.SetCellFormula(17, 2, "=SUM(B9:E9)+G9")
                                            'outputExcel.SetCellFormula(17, 3, "=SUM(F5:F8)")
                                            'outputExcel.SetCellFormula(17, 4, "=B17/$B$19")
                                            'outputExcel.SetCellFormula(17, 5, "=C17/$B$19")

                                            'outputExcel.SetData(18, 1, "Total Changes")
                                            'outputExcel.SetCellFormula(18, 2, "=SUM(B14:B17)")
                                            'outputExcel.SetCellFormula(18, 3, "=SUM(C14:C17)")
                                            'outputExcel.SetCellFormula(18, 4, "=B18/$B$19")
                                            'outputExcel.SetCellFormula(18, 5, "=C18/$B$19")

                                            'outputExcel.SetData(19, 1, "Total Headcount")
                                            'outputExcel.SetCellFormula(19, 2, "=H10")

                                            outputExcel.SaveExcel()
                                        End Using
                                        Dim outputFile As String = Path.Combine(_outputDirectoryName, Path.GetFileName(skillOutputFileName))
                                        If File.Exists(outputFile) Then File.Delete(outputFile)
                                        File.Copy(skillOutputFileName, outputFile)
                                    End If
                                Else
                                    If scoreFilename Is Nothing Then OnHeartbeatError(String.Format("Score File not exists for {0}", skillName))
                                End If
                            End If
                        Next
                    End If
                End Using
            Else
                Throw New ApplicationException("Mapping file missing")
            End If
        Catch aex As Exception
            Throw aex
        End Try
    End Function

    Private Function GetSheetName(ByVal allSheets As List(Of String), ByVal subskill As String) As String
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
        If allSheets IsNot Nothing Then
            While allSheets.Contains(ret, StringComparer.OrdinalIgnoreCase)
                ret = String.Format("{0}_1", ret)
            End While
        End If
        Return ret
    End Function

    Private Sub SetDataTo2DArray(ByVal rowNumber As Integer, ByVal columnNumber As Integer, ByVal data As String, ByRef array As Object(,))
        array(rowNumber - 9, columnNumber - 10) = data
    End Sub

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