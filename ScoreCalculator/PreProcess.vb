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
    Private ReadOnly monthList As Dictionary(Of String, String)
    Private ReadOnly scoreFileSchema As Dictionary(Of String, String)
    Private ReadOnly empFileSchema As Dictionary(Of String, String)

    Public Sub New(ByVal canceller As CancellationTokenSource, ByVal directoryName As String, ByVal mappingFile As String)
        _cts = canceller
        _inputDirectoryName = directoryName
        _outputDirectoryName = Path.Combine(Directory.GetParent(_inputDirectoryName).FullName, "In Process")
        _mappingFile = mappingFile
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
        OnHeartbeat("Opening mapping file")
        Using mappingXL As New ExcelHelper(_mappingFile, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
            AddHandler mappingXL.Heartbeat, AddressOf OnHeartbeat
            AddHandler mappingXL.WaitingFor, AddressOf OnWaitingFor
            Dim mappingSheetList As List(Of String) = mappingXL.GetExcelSheetsName()
            Dim skillList As List(Of String) = Nothing
            For Each runningFile In Directory.GetFiles(_inputDirectoryName)
                If runningFile.ToUpper.Contains("BFSI") Then
                    Continue For
                End If
                If skillList Is Nothing OrElse Not skillList.Contains(runningFile) Then
                    Dim pairFound As Boolean = False
                    For Each runningPairFile In Directory.GetFiles(_inputDirectoryName)
                        If runningPairFile.ToUpper <> runningFile.ToUpper AndAlso
                            runningPairFile.Remove(runningPairFile.Count - 13) = runningFile.Remove(runningFile.Count - 13) Then
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
                                Dim skillName As String = Path.GetFileName(currentFileName).Remove(Path.GetFileName(currentFileName).Count - 29)
                                Dim foundationSkillList As List(Of String) = Nothing
                                If mappingSheetList IsNot Nothing AndAlso mappingSheetList.Count > 0 Then
                                    Dim skillMappingSheetName As String = Nothing
                                    For Each sheet In mappingSheetList
                                        If sheet.ToUpper.Contains(skillName.ToUpper) Then
                                            skillMappingSheetName = sheet
                                            Exit For
                                        End If
                                    Next
                                    If skillMappingSheetName IsNot Nothing Then
                                        mappingXL.SetActiveSheet(skillMappingSheetName)
                                        Dim mappingData As Object(,) = mappingXL.GetExcelInMemory()
                                        foundationSkillList = GetFoundationList(mappingData)
                                    End If
                                End If
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
                                    For rowCounter As Integer = 2 To currentMonthScoreData.GetLength(0) - 1
                                        OnHeartbeat(String.Format("Replacing Zero score {0}/{1}. File:{2}", rowCounter, currentMonthScoreData.GetLength(0) - 1, currentFileName))
                                        _cts.Token.ThrowIfCancellationRequested()
                                        Dim empID As String = currentMonthScoreData(rowCounter, 1)
                                        If empID IsNot Nothing Then
                                            Dim previousScoreRow As Integer = _cmn.GetRowOf2DArray(previousMonthScoreData, 1, empID, True)
                                            If previousScoreRow <> Integer.MinValue Then
                                                For columnCounter As Integer = 3 To currentMonthScoreData.GetLength(1) - 1
                                                    _cts.Token.ThrowIfCancellationRequested()
                                                    Dim score As String = currentMonthScoreData(rowCounter, columnCounter)
                                                    If score IsNot Nothing AndAlso score = 0 Then
                                                        Dim columnName As String = currentMonthScoreData(1, columnCounter)
                                                        Dim previousScoreColumn As Integer = _cmn.GetColumnOf2DArray(previousMonthScoreData, 1, columnName)
                                                        If previousScoreColumn <> Integer.MinValue Then
                                                            currentMonthScoreData(rowCounter, columnCounter) = previousMonthScoreData(previousScoreRow, previousScoreColumn)
                                                        End If
                                                    ElseIf score IsNot Nothing Then
                                                        Dim columnName As String = currentMonthScoreData(1, columnCounter)
                                                        If foundationSkillList IsNot Nothing AndAlso foundationSkillList.Count > 0 AndAlso
                                                            foundationSkillList.Contains(columnName, StringComparer.OrdinalIgnoreCase) Then
                                                            Dim previousScoreColumn As Integer = _cmn.GetColumnOf2DArray(previousMonthScoreData, 1, columnName)
                                                            If previousScoreColumn <> Integer.MinValue Then
                                                                Dim previousScore As String = previousMonthScoreData(previousScoreRow, previousScoreColumn)
                                                                If previousScore IsNot Nothing AndAlso Val(previousScore) > Val(score) Then
                                                                    currentMonthScoreData(rowCounter, columnCounter) = Val(previousScore)
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Next
                                            End If
                                        End If
                                    Next
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
        End Using
    End Function

    Public Async Function ProcessEmployeeData() As Task
        Await Task.Delay(1).ConfigureAwait(False)
        If Not Directory.Exists(_outputDirectoryName) Then
            Throw New ApplicationException(String.Format("Directory not found. {0}", _outputDirectoryName))
        End If
        Dim skillList As List(Of String) = Nothing
        For Each runningFile In Directory.GetFiles(_inputDirectoryName)
            If runningFile.ToUpper.Contains("BFSI") AndAlso (skillList Is Nothing OrElse Not skillList.Contains(runningFile)) Then
                Dim pairFound As Boolean = False
                For Each runningPairFile In Directory.GetFiles(_inputDirectoryName)
                    If runningPairFile.ToUpper <> runningFile.ToUpper AndAlso
                        runningPairFile.Remove(runningPairFile.Count - 13) = runningFile.Remove(runningFile.Count - 13) Then
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
