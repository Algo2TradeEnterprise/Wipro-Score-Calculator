Imports System.IO
Imports Utilities.DAL
Imports System.Threading

Public Class ScoreModifier
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
    Private Class EmployeeDetails
        Public EmpID As String
        Public Practice As String
        Public Sheet As String
    End Class

    Private Class WeightageDetails
        Public Bucket As String
        Public Weightage As Decimal
        Public ColumnNumber As Integer
    End Class
#End Region

    Private ReadOnly _cts As CancellationTokenSource
    Private ReadOnly _directoryName As String
    Private ReadOnly _employeeListFileName As String
    Private ReadOnly _outputDirectoryName As String

    Private ReadOnly _ModifyPendingToComplete As Boolean
    Private ReadOnly _ModifyPendingToCompleteMinScore As Decimal
    Private ReadOnly _ModifyCompleteToITPi As Boolean
    Private ReadOnly _ModifyCompleteToITPiMinScore As Decimal

    Public Sub New(ByVal canceller As CancellationTokenSource, ByVal directoryPath As String,
                   ByVal mdfyPndngToCmplt As Boolean, ByVal mdfyCmpltToITPi As Boolean,
                   ByVal pndngToCmpltMinScore As Decimal, ByVal cmpltToITPiMinScore As Decimal)
        _cts = canceller
        _directoryName = directoryPath
        _ModifyPendingToComplete = mdfyPndngToCmplt
        _ModifyPendingToCompleteMinScore = pndngToCmpltMinScore
        _ModifyCompleteToITPi = mdfyCmpltToITPi
        _ModifyCompleteToITPiMinScore = cmpltToITPiMinScore

        _employeeListFileName = Path.Combine(My.Application.Info.DirectoryPath, "Excel Test", "Pre Process", "Score Modifier", "List.csv")

        _outputDirectoryName = Path.Combine(My.Application.Info.DirectoryPath, "Excel Test", "Pre Process")
    End Sub

    Public Async Function ProcessDataAsync() As Task
        Await Task.Delay(1).ConfigureAwait(False)

        Dim outputFiles As List(Of String) = Nothing
        Dim rawScoreFiles As List(Of String) = Nothing
        For Each runningFile In Directory.GetFiles(_directoryName)
            Dim filename As String = Path.GetFileName(runningFile)
            Dim extension As String = Path.GetExtension(runningFile)
            If extension.ToUpper.Contains("XLS") Then
                If filename.ToUpper.Contains("OUTPUT") Then
                    If outputFiles Is Nothing Then outputFiles = New List(Of String)
                    outputFiles.Add(runningFile)
                Else
                    If rawScoreFiles Is Nothing Then rawScoreFiles = New List(Of String)
                    rawScoreFiles.Add(runningFile)
                End If
            End If
        Next
        If outputFiles IsNot Nothing AndAlso outputFiles.Count > 0 AndAlso
           rawScoreFiles IsNot Nothing AndAlso rawScoreFiles.Count > 0 Then
            Dim practiceList As List(Of String) = Nothing
            Dim empDetailsList As List(Of EmployeeDetails) = Nothing
            Dim prctCtr As Integer = 0
            For Each runningFile In outputFiles
                Dim filePractice As String = Path.GetFileNameWithoutExtension(runningFile).Trim.Split(" ")(0).Trim
                prctCtr += 1
                OnHeartbeat(String.Format("Creating employee list from previous month output to update score for prectice:{0} #{1}/{2}", filePractice, prctCtr, outputFiles.Count))
                Using runningXL As New ExcelHelper(runningFile, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                    runningXL.SetActiveSheet("I T Pi")

                    Dim range As String = runningXL.GetNamedRange(10, 0, 1, runningXL.GetLastCol(10))
                    Dim statusColumnNumberList As List(Of KeyValuePair(Of Integer, Integer)) = runningXL.FindAll("With Foundation", range, True)
                    Dim scoreColumnNumberList As List(Of KeyValuePair(Of Integer, Integer)) = runningXL.FindAll("Foundation Level Proficiency", range, True)

                    If scoreColumnNumberList IsNot Nothing AndAlso scoreColumnNumberList.Count > 0 AndAlso
                        statusColumnNumberList IsNot Nothing AndAlso statusColumnNumberList.Count > 0 Then
                        Dim statusColumnNumber As Integer = statusColumnNumberList.FirstOrDefault.Value
                        Dim scoreColumnNumber As Integer = scoreColumnNumberList.FirstOrDefault.Value

                        Dim dataRange As String = runningXL.GetNamedRange(11, runningXL.GetLastRow(1) - 1, 1, runningXL.GetLastCol(11) - 1)
                        Dim data As Object(,) = runningXL.GetExcelInMemory(dataRange)

                        For rowCtr As Integer = 2 To data.GetLength(0)
                            Dim empID As String = data(rowCtr, 1)
                            If empID IsNot Nothing Then
                                Dim withFoundationStatus As String = data(rowCtr, statusColumnNumber)
                                If withFoundationStatus IsNot Nothing Then
                                    If withFoundationStatus.ToUpper = "FOUNDATION PENDING" Then
                                        'Dim withoutFoundationStatus As String = data(rowCtr, statusColumnNumber + 1)
                                        'If withoutFoundationStatus.ToUpper = "I" OrElse
                                        '    withoutFoundationStatus.ToUpper = "T" OrElse
                                        '    withoutFoundationStatus.ToUpper = "PI" Then
                                        Dim foundationScore As String = data(rowCtr, scoreColumnNumber)
                                        If foundationScore IsNot Nothing AndAlso IsNumeric(foundationScore) Then
                                            If Val(foundationScore) >= _ModifyPendingToCompleteMinScore Then
                                                Dim empDetails As EmployeeDetails = New EmployeeDetails With {
                                                                                        .EmpID = empID,
                                                                                        .Practice = filePractice,
                                                                                        .Sheet = "FOUNDATION"
                                                                                    }

                                                If empDetailsList Is Nothing Then empDetailsList = New List(Of EmployeeDetails)
                                                empDetailsList.Add(empDetails)
                                            End If
                                        End If
                                        'End If
                                    ElseIf withFoundationStatus.ToUpper = "FOUNDATION COMPLETE" Then
                                        Dim empDetails As EmployeeDetails = New EmployeeDetails With {
                                                                                            .EmpID = empID,
                                                                                            .Practice = filePractice,
                                                                                            .Sheet = "I T PI"
                                                                                        }

                                        If empDetailsList Is Nothing Then empDetailsList = New List(Of EmployeeDetails)
                                        empDetailsList.Add(empDetails)
                                    End If
                                End If
                            End If
                        Next
                    End If
                End Using

                If practiceList Is Nothing Then practiceList = New List(Of String)
                practiceList.Add(filePractice.ToUpper)
            Next

            If practiceList IsNot Nothing AndAlso practiceList.Count > 0 AndAlso
                empDetailsList IsNot Nothing AndAlso empDetailsList.Count > 0 Then
                Dim practiceCtr As Integer = 0
                For Each runningPractice In practiceList
                    practiceCtr += 1
                    Dim currentPracticeOutputFile As String = GetCurrentPracticeOutputFile(outputFiles, runningPractice)
                    Dim currentPracticeScoreFile As String = GetCurrentPracticeScoreFile(rawScoreFiles, runningPractice)
                    If currentPracticeOutputFile Is Nothing Then
                        OnHeartbeatError(String.Format("Output File not available for practice: {0}", runningPractice))
                    End If
                    If currentPracticeScoreFile Is Nothing Then
                        OnHeartbeatError(String.Format("Score File not available for practice: {0}", runningPractice))
                    End If
                    If currentPracticeScoreFile IsNot Nothing AndAlso currentPracticeOutputFile IsNot Nothing Then
                        OnHeartbeat(String.Format("Opening output file {0}", currentPracticeOutputFile))
                        If _ModifyPendingToComplete Then
                            Using outputXL As New ExcelHelper(currentPracticeOutputFile, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                                outputXL.SetActiveSheet("Foundation")

                                OnHeartbeat(String.Format("Getting Weightage Data for {0} #Pass 1", runningPractice))
                                Dim weightageData As List(Of WeightageDetails) = Nothing
                                Dim lastWeightageColumn As Integer = outputXL.GetLastCol(9)
                                For columnCtr As Integer = 1 To lastWeightageColumn
                                    Dim weightage As String = outputXL.GetData(9, columnCtr)
                                    If weightage IsNot Nothing AndAlso weightage <> "" AndAlso IsNumeric(weightage) Then
                                        Dim bucket As String = outputXL.GetData(10, columnCtr)

                                        Dim wtgDtls As WeightageDetails = New WeightageDetails With {
                                    .Weightage = CDec(weightage) * 100,
                                    .Bucket = bucket,
                                    .ColumnNumber = columnCtr
                                }

                                        If weightageData Is Nothing Then weightageData = New List(Of WeightageDetails)
                                        weightageData.Add(wtgDtls)
                                    End If
                                Next

                                If weightageData IsNot Nothing AndAlso weightageData.Count > 0 Then
                                    OnHeartbeat(String.Format("Opening score file {0} #Pass 1", currentPracticeScoreFile))
                                    Using scoreXL As New ExcelHelper(currentPracticeScoreFile, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                                        OnHeartbeat(String.Format("Processing for {0} #{1}/{2} #Pass 1", runningPractice, practiceCtr, practiceList.Count))
                                        Dim practiceEmpDetailsList As IEnumerable(Of EmployeeDetails) = empDetailsList.Where(Function(x)
                                                                                                                                 Return x.Practice.ToUpper = runningPractice.ToUpper AndAlso
                                                                                                                         x.Sheet.ToUpper = "FOUNDATION".ToUpper
                                                                                                                             End Function)
                                        If practiceEmpDetailsList IsNot Nothing AndAlso practiceEmpDetailsList.Count > 0 Then
                                            Dim empCtr As Integer = 0
                                            For Each runningEmp In practiceEmpDetailsList
                                                If runningEmp.Practice.ToUpper = runningPractice.ToUpper Then
                                                    empCtr += 1
                                                    OnHeartbeat(String.Format("Processing for {0} #{1}/{2}, #{3}/{4} #Pass 1", runningPractice, practiceCtr, practiceList.Count, empCtr, practiceEmpDetailsList.Count))

                                                    Dim range As String = outputXL.GetNamedRange(12, outputXL.GetLastRow(1), 1, 1)
                                                    Dim empRowsColumns As List(Of KeyValuePair(Of Integer, Integer)) = outputXL.FindAll(runningEmp.EmpID, range, True)
                                                    If empRowsColumns IsNot Nothing AndAlso empRowsColumns.Count > 0 Then
                                                        Dim empRowInOutputFile As Integer = empRowsColumns.FirstOrDefault.Key

                                                        Dim scoreRange As String = scoreXL.GetNamedRange(2, scoreXL.GetLastRow(1), 1, 1)
                                                        Dim empRowsColumnsScoreFile As List(Of KeyValuePair(Of Integer, Integer)) = scoreXL.FindAll(runningEmp.EmpID, scoreRange, True)
                                                        If empRowsColumnsScoreFile IsNot Nothing AndAlso empRowsColumnsScoreFile.Count > 0 Then
                                                            Dim empRowInScore As Integer = empRowsColumnsScoreFile.FirstOrDefault.Key

                                                            Dim wtgCtr As Integer = 0
                                                            For Each runningWeightage In weightageData.OrderByDescending(Function(x)
                                                                                                                             Return x.Weightage
                                                                                                                         End Function)
                                                                outputXL.SetActiveSheet("Foundation")
                                                                Dim skillRowNumber As Integer = 11
                                                                Dim startingColumnNumber As Integer = runningWeightage.ColumnNumber
                                                                Dim enndingColumnNumber As Integer = startingColumnNumber + 1
                                                                'Finding next 'BEST SCORE FROM Group' column
                                                                For colCtr As Integer = startingColumnNumber To outputXL.GetLastCol(skillRowNumber)
                                                                    Dim columnName As String = outputXL.GetData(skillRowNumber, colCtr)
                                                                    If columnName.ToUpper = "BEST SCORE FROM GROUP" Then
                                                                        enndingColumnNumber = colCtr - 1
                                                                        Exit For
                                                                    End If
                                                                Next

                                                                Dim lastHighestScore As Decimal = Decimal.MinValue
                                                                Dim lastHighestScoreColumnName As String = Nothing
                                                                Dim lastHighestScoreColumnNumber As Integer = Integer.MinValue
                                                                For colCtr As Integer = startingColumnNumber To enndingColumnNumber
                                                                    Dim score As String = outputXL.GetData(empRowInOutputFile, colCtr)
                                                                    If score IsNot Nothing AndAlso IsNumeric(score) Then
                                                                        If CDec(score) > lastHighestScore Then
                                                                            lastHighestScore = CDec(score)
                                                                            lastHighestScoreColumnName = outputXL.GetData(skillRowNumber, colCtr)
                                                                            lastHighestScoreColumnNumber = colCtr
                                                                        End If
                                                                    End If
                                                                Next

                                                                If lastHighestScoreColumnName IsNot Nothing AndAlso lastHighestScoreColumnNumber <> Integer.MinValue Then
                                                                    'Getting duplicate column count
                                                                    Dim totalDuplicateColumnCount As Integer = 0
                                                                    For colCtr As Integer = lastHighestScoreColumnNumber To 1 Step -1
                                                                        Dim columnName As String = outputXL.GetData(skillRowNumber, colCtr)
                                                                        If columnName IsNot Nothing AndAlso columnName.ToUpper = lastHighestScoreColumnName.ToUpper Then
                                                                            totalDuplicateColumnCount += 1
                                                                        End If
                                                                    Next
                                                                    Dim sheetName As String = GetSheetName(lastHighestScoreColumnName)
                                                                    If totalDuplicateColumnCount > 1 Then
                                                                        For ctr As Integer = 1 To totalDuplicateColumnCount - 1
                                                                            sheetName = String.Format("{0}_1", sheetName)
                                                                        Next
                                                                    End If
                                                                    If outputXL.GetExcelSheetsName().Contains(sheetName) Then
                                                                        outputXL.SetActiveSheet(sheetName)
                                                                        Dim maxScoreRange As String = outputXL.GetNamedRange(2, outputXL.GetLastRow(1), 1, 1)
                                                                        Dim empRowsColumnsMaxScoreSheet As List(Of KeyValuePair(Of Integer, Integer)) = outputXL.FindAll(runningEmp.EmpID, maxScoreRange, True)
                                                                        If empRowsColumnsMaxScoreSheet IsNot Nothing AndAlso empRowsColumnsMaxScoreSheet.Count > 0 Then
                                                                            Dim empRowInMaxScore As Integer = empRowsColumnsMaxScoreSheet.FirstOrDefault.Key

                                                                            Dim lastHighestMaxScore As Decimal = Decimal.MinValue
                                                                            Dim lastHighestMaxScoreColumnName As String = Nothing
                                                                            For colCtr As Integer = 3 To outputXL.GetLastCol(empRowInMaxScore)
                                                                                Dim score As String = outputXL.GetData(empRowInMaxScore, colCtr)
                                                                                If score IsNot Nothing AndAlso IsNumeric(score) Then
                                                                                    If CDec(score) > lastHighestMaxScore Then
                                                                                        lastHighestMaxScore = CDec(score)
                                                                                        lastHighestMaxScoreColumnName = outputXL.GetData(1, colCtr)
                                                                                    End If
                                                                                End If
                                                                            Next
                                                                            If lastHighestMaxScoreColumnName IsNot Nothing Then
                                                                                For colctr As Integer = 1 To scoreXL.GetLastCol(1)
                                                                                    Dim columnName As String = scoreXL.GetData(1, colctr)
                                                                                    If columnName IsNot Nothing AndAlso columnName.ToUpper = lastHighestMaxScoreColumnName.ToUpper Then
                                                                                        Dim score As String = scoreXL.GetData(empRowInScore, colctr)
                                                                                        Dim currentScore As Decimal = 0
                                                                                        If score IsNot Nothing AndAlso IsNumeric(score) Then
                                                                                            currentScore = score
                                                                                        End If
                                                                                        If lastHighestMaxScore = 0 Then lastHighestMaxScore = 30
                                                                                        Dim projectedScore As Decimal = lastHighestMaxScore + lastHighestMaxScore / 2 + 0.1
                                                                                        If projectedScore > currentScore Then
                                                                                            If projectedScore > 100 Then projectedScore = 100

                                                                                            scoreXL.SetData(empRowInScore, colctr, projectedScore, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                                                                            Exit For
                                                                                        End If
                                                                                    End If
                                                                                Next
                                                                            Else
                                                                                'OnHeartbeatError(String.Format("Neglected Emp ID:{0}, Practice:{1}, Reason:{2}",
                                                                                '  runningEmp.EmpID, runningPractice, "Highest Score column name Not found in 'Max Score' sheet"))
                                                                            End If
                                                                        Else
                                                                            'OnHeartbeatError(String.Format("Neglected Emp ID:{0}, Practice:{1}, Reason:{2}",
                                                                            '      runningEmp.EmpID, runningPractice, "EMP ID Not found in 'Max Score' sheet"))
                                                                        End If
                                                                    Else
                                                                        'OnHeartbeatError(String.Format("Neglected Emp ID:{0}, Practice:{1}, Reason:'{2}'Sheet Not found",
                                                                        '      runningEmp.EmpID, runningPractice, sheetName))
                                                                    End If
                                                                Else
                                                                    'OnHeartbeatError(String.Format("Neglected Emp ID:{0}, Practice:{1}, Reason:{2}",
                                                                    '          runningEmp.EmpID, runningPractice, "Highest Score column name Not found in 'Foundation' sheet"))
                                                                End If

                                                                wtgCtr += 1
                                                                If wtgCtr = 2 Then Exit For
                                                            Next
                                                        Else
                                                            OnHeartbeatError(String.Format("Neglected Emp ID:{0}, Practice:{1}, Reason:{2}",
                                                                             runningEmp.EmpID, runningPractice, "EMP ID Not found in 'Raw Score' file"))
                                                        End If
                                                    Else
                                                        OnHeartbeatError(String.Format("Neglected Emp ID:{0}, Practice:{1}, Reason:{2}",
                                                                               runningEmp.EmpID, runningPractice, "EMP ID Not found in 'Foundation' sheet"))
                                                    End If
                                                End If
                                            Next
                                        End If
                                    End Using
                                End If
                            End Using
                        End If
                        If _ModifyCompleteToITPi Then
                            'Working on I T Pi
                            Using outputXL As New ExcelHelper(currentPracticeOutputFile, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                                outputXL.SetActiveSheet("I T Pi")

                                OnHeartbeat(String.Format("Getting Score Columns for {0} #Pass 2", runningPractice))
                                Dim scoreStartingColumn As Integer = Integer.MinValue
                                Dim scoreEndingColumn As Integer = Integer.MinValue
                                For columnCtr As Integer = 1 To outputXL.GetLastCol(10)
                                    Dim bucket As String = outputXL.GetData(10, columnCtr)
                                    If bucket IsNot Nothing AndAlso bucket <> "" Then
                                        If scoreStartingColumn = Integer.MinValue Then scoreStartingColumn = columnCtr
                                        If bucket.ToUpper = "Foundation Level Proficiency".ToUpper Then
                                            scoreEndingColumn = columnCtr - 1
                                            Exit For
                                        End If
                                    End If
                                Next

                                If scoreStartingColumn <> Integer.MinValue AndAlso scoreEndingColumn <> Integer.MinValue Then
                                    OnHeartbeat(String.Format("Opening score file {0} #Pass 2", currentPracticeScoreFile))
                                    Using scoreXL As New ExcelHelper(currentPracticeScoreFile, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                                        OnHeartbeat(String.Format("Processing for {0} #{1}/{2} #Pass 2", runningPractice, practiceCtr, practiceList.Count))
                                        Dim practiceEmpDetailsList As IEnumerable(Of EmployeeDetails) = empDetailsList.Where(Function(x)
                                                                                                                                 Return x.Practice.ToUpper = runningPractice.ToUpper AndAlso
                                                                                                                         x.Sheet.ToUpper = "I T PI".ToUpper
                                                                                                                             End Function)
                                        If practiceEmpDetailsList IsNot Nothing AndAlso practiceEmpDetailsList.Count > 0 Then
                                            Dim empCtr As Integer = 0
                                            For Each runningEmp In practiceEmpDetailsList
                                                If runningEmp.Practice.ToUpper = runningPractice.ToUpper Then
                                                    empCtr += 1
                                                    OnHeartbeat(String.Format("Processing for {0} #{1}/{2}, #{3}/{4} #Pass 2", runningPractice, practiceCtr, practiceList.Count, empCtr, practiceEmpDetailsList.Count))

                                                    Dim range As String = outputXL.GetNamedRange(12, outputXL.GetLastRow(1), 1, 1)
                                                    Dim empRowsColumns As List(Of KeyValuePair(Of Integer, Integer)) = outputXL.FindAll(runningEmp.EmpID, range, True)
                                                    If empRowsColumns IsNot Nothing AndAlso empRowsColumns.Count > 0 Then
                                                        Dim empRowInOutputFile As Integer = empRowsColumns.FirstOrDefault.Key

                                                        Dim scoreRange As String = scoreXL.GetNamedRange(2, scoreXL.GetLastRow(1), 1, 1)
                                                        Dim empRowsColumnsScoreFile As List(Of KeyValuePair(Of Integer, Integer)) = scoreXL.FindAll(runningEmp.EmpID, scoreRange, True)
                                                        If empRowsColumnsScoreFile IsNot Nothing AndAlso empRowsColumnsScoreFile.Count > 0 Then
                                                            Dim empRowInScore As Integer = empRowsColumnsScoreFile.FirstOrDefault.Key

                                                            outputXL.SetActiveSheet("I T Pi")
                                                            Dim skillRowNumber As Integer = 11
                                                            Dim lastHighestScoreColumnName As String = Nothing
                                                            Dim lastHighestScoreColumnNumber As Integer = Integer.MinValue
                                                            For colCtr As Integer = scoreStartingColumn To scoreEndingColumn
                                                                Dim score As String = outputXL.GetData(empRowInOutputFile, colCtr)
                                                                If score IsNot Nothing AndAlso IsNumeric(score) Then
                                                                    If CDec(score) >= _ModifyCompleteToITPiMinScore Then
                                                                        lastHighestScoreColumnName = outputXL.GetData(skillRowNumber, colCtr)
                                                                        lastHighestScoreColumnNumber = colCtr
                                                                        Exit For
                                                                    End If
                                                                End If
                                                            Next

                                                            If lastHighestScoreColumnName IsNot Nothing AndAlso lastHighestScoreColumnNumber <> Integer.MinValue Then
                                                                'Getting duplicate column count
                                                                Dim totalDuplicateColumnCount As Integer = 0
                                                                outputXL.SetActiveSheet("Foundation")
                                                                For colCtr As Integer = 1 To outputXL.GetLastCol(skillRowNumber)
                                                                    Dim columnName As String = outputXL.GetData(skillRowNumber, colCtr)
                                                                    If columnName IsNot Nothing AndAlso columnName.ToUpper = lastHighestScoreColumnName.ToUpper Then
                                                                        totalDuplicateColumnCount += 1
                                                                    End If
                                                                Next
                                                                outputXL.SetActiveSheet("I T Pi")
                                                                For colCtr As Integer = lastHighestScoreColumnNumber To 1 Step -1
                                                                    Dim columnName As String = outputXL.GetData(skillRowNumber, colCtr)
                                                                    If columnName IsNot Nothing AndAlso columnName.ToUpper = lastHighestScoreColumnName.ToUpper Then
                                                                        totalDuplicateColumnCount += 1
                                                                    End If
                                                                Next
                                                                Dim sheetName As String = GetSheetName(lastHighestScoreColumnName)
                                                                If totalDuplicateColumnCount > 1 Then
                                                                    For ctr As Integer = 1 To totalDuplicateColumnCount - 1
                                                                        sheetName = String.Format("{0}_1", sheetName)
                                                                    Next
                                                                End If
                                                                If outputXL.GetExcelSheetsName().Contains(sheetName) Then
                                                                    outputXL.SetActiveSheet(sheetName)
                                                                    Dim maxScoreRange As String = outputXL.GetNamedRange(2, outputXL.GetLastRow(1), 1, 1)
                                                                    Dim empRowsColumnsMaxScoreSheet As List(Of KeyValuePair(Of Integer, Integer)) = outputXL.FindAll(runningEmp.EmpID, maxScoreRange, True)
                                                                    If empRowsColumnsMaxScoreSheet IsNot Nothing AndAlso empRowsColumnsMaxScoreSheet.Count > 0 Then
                                                                        Dim empRowInMaxScore As Integer = empRowsColumnsMaxScoreSheet.FirstOrDefault.Key

                                                                        Dim lastHighestMaxScore As Decimal = Decimal.MinValue
                                                                        Dim lastHighestMaxScoreColumnName As String = Nothing
                                                                        For colCtr As Integer = 3 To outputXL.GetLastCol(empRowInMaxScore)
                                                                            Dim score As String = outputXL.GetData(empRowInMaxScore, colCtr)
                                                                            If score IsNot Nothing AndAlso IsNumeric(score) Then
                                                                                If CDec(score) > lastHighestMaxScore Then
                                                                                    lastHighestMaxScore = CDec(score)
                                                                                    lastHighestMaxScoreColumnName = outputXL.GetData(1, colCtr)
                                                                                End If
                                                                            End If
                                                                        Next
                                                                        If lastHighestMaxScoreColumnName IsNot Nothing Then
                                                                            For colctr As Integer = 1 To scoreXL.GetLastCol(1)
                                                                                Dim columnName As String = scoreXL.GetData(1, colctr)
                                                                                If columnName IsNot Nothing AndAlso columnName.ToUpper = lastHighestMaxScoreColumnName.ToUpper Then
                                                                                    Dim score As String = scoreXL.GetData(empRowInScore, colctr)
                                                                                    Dim currentScore As Decimal = 0
                                                                                    If score IsNot Nothing AndAlso IsNumeric(score) Then
                                                                                        currentScore = score
                                                                                    End If
                                                                                    If lastHighestMaxScore = 0 Then lastHighestMaxScore = 50
                                                                                    Dim projectedScore As Decimal = lastHighestMaxScore + lastHighestMaxScore / 2 + 0.1
                                                                                    If projectedScore > currentScore Then
                                                                                        If projectedScore > 100 Then projectedScore = 100

                                                                                        scoreXL.SetData(empRowInScore, colctr, projectedScore, "##,##,##0.00", ExcelHelper.XLAlign.Right)
                                                                                        Exit For
                                                                                    End If
                                                                                End If
                                                                            Next
                                                                        Else
                                                                            'OnHeartbeatError(String.Format("Neglected Emp ID:{0}, Practice:{1}, Reason:{2}",
                                                                            '  runningEmp.EmpID, runningPractice, "Highest Score column name Not found in 'Max Score' sheet"))
                                                                        End If
                                                                    Else
                                                                        'OnHeartbeatError(String.Format("Neglected Emp ID:{0}, Practice:{1}, Reason:{2}",
                                                                        '      runningEmp.EmpID, runningPractice, "EMP ID Not found in 'Max Score' sheet"))
                                                                    End If
                                                                Else
                                                                    'OnHeartbeatError(String.Format("Neglected Emp ID:{0}, Practice:{1}, Reason:'{2}'Sheet Not found",
                                                                    '      runningEmp.EmpID, runningPractice, sheetName))
                                                                End If
                                                            Else
                                                                'OnHeartbeatError(String.Format("Neglected Emp ID:{0}, Practice:{1}, Reason:{2}",
                                                                '          runningEmp.EmpID, runningPractice, "Above min Score Not found in 'I T Pi' sheet"))
                                                            End If
                                                        Else
                                                            OnHeartbeatError(String.Format("Neglected Emp ID:{0}, Practice:{1}, Reason:{2}",
                                                                         runningEmp.EmpID, runningPractice, "EMP ID Not found in 'Raw Score' file"))
                                                        End If
                                                    Else
                                                        OnHeartbeatError(String.Format("Neglected Emp ID:{0}, Practice:{1}, Reason:{2}",
                                                                           runningEmp.EmpID, runningPractice, "EMP ID Not found in 'I T Pi' sheet"))
                                                    End If
                                                End If
                                            Next
                                        End If
                                    End Using
                                End If
                            End Using
                        End If
                        Dim outputFile As String = Path.Combine(_outputDirectoryName, Path.GetFileName(currentPracticeScoreFile))
                        If File.Exists(outputFile) Then File.Delete(outputFile)
                        File.Copy(currentPracticeScoreFile, outputFile)
                    End If
                Next
            End If
        End If
    End Function

    Private Function GetCurrentPracticeScoreFile(ByVal scoreFilesList As List(Of String), ByVal practice As String) As String
        Dim ret As String = Nothing
        If scoreFilesList IsNot Nothing AndAlso scoreFilesList.Count > 0 Then
            For Each runningFile In scoreFilesList
                Dim filePractice As String = Path.GetFileNameWithoutExtension(runningFile).Trim.Split("_")(0).Trim
                If filePractice.ToUpper = practice.ToUpper Then
                    ret = runningFile
                    Exit For
                End If
            Next
        End If
        Return ret
    End Function

    Private Function GetCurrentPracticeOutputFile(ByVal outputFilesList As List(Of String), ByVal practice As String) As String
        Dim ret As String = Nothing
        If outputFilesList IsNot Nothing AndAlso outputFilesList.Count > 0 Then
            For Each runningFile In outputFilesList
                Dim filePractice As String = Path.GetFileNameWithoutExtension(runningFile).Trim.Split(" ")(0).Trim
                If filePractice.ToUpper = practice.ToUpper Then
                    ret = runningFile
                    Exit For
                End If
            Next
        End If
        Return ret
    End Function

    Private Function GetSheetName(ByVal subskill As String) As String
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
