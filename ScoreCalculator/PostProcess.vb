Imports System.IO
Imports Utilities.DAL
Imports System.Threading

Public Class PostProcess
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
    Private ReadOnly _employeeFile As String
    Private ReadOnly _directoryName As String
    Private ReadOnly _quarterlyProgress As Dictionary(Of String, String)
    Public Sub New(ByVal canceller As CancellationTokenSource, ByVal employeeFile As String)
        _cts = canceller
        _employeeFile = employeeFile
        _directoryName = Path.GetDirectoryName(_employeeFile)
        _quarterlyProgress = New Dictionary(Of String, String) From
            {{"JAN", "Q4"},
             {"FEB", "Q4"},
             {"MAR", "Q4"},
             {"APR", "Q1"},
             {"MAY", "Q1"},
             {"JUN", "Q1"},
             {"JUL", "Q2"},
             {"AUG", "Q2"},
             {"SEP", "Q2"},
             {"OCT", "Q3"},
             {"NOV", "Q3"},
             {"DEC", "Q3"}}
    End Sub

    Public Async Function ProcessData() As Task
        Await Task.Delay(1).ConfigureAwait(False)
        Dim jscoreList As Dictionary(Of String, JScore) = Nothing
        For Each runningFile In Directory.GetFiles(_directoryName)
            If runningFile.Contains("Output") Then
                OnHeartbeat(String.Format("Opening {0}", runningFile))
                Dim practice As String = Path.GetFileNameWithoutExtension(runningFile).Trim.Split(" ")(0)
                Using xl As New ExcelHelper(runningFile, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                    AddHandler xl.Heartbeat, AddressOf OnHeartbeat
                    AddHandler xl.WaitingFor, AddressOf OnWaitingFor

                    xl.SetActiveSheet("I T Pi")
                    Dim lastRow As Integer = xl.GetLastRow(1)
                    Dim dataRange As String = xl.GetNamedRange(10, 10, 1, xl.GetLastCol(11) - 1)
                    Dim foundationColumnNumber As Integer = xl.FindAll("With Foundation", dataRange, True).FirstOrDefault.Value
                    For rowCounter As Integer = 12 To lastRow
                        OnHeartbeat(String.Format("Reading {0}/{1} of {2}", rowCounter, lastRow, runningFile))
                        Dim jscr As JScore = New JScore With {
                            .EmpID = xl.GetData(rowCounter, 1),
                            .WithFoundationITPi = xl.GetData(rowCounter, foundationColumnNumber),
                            .WithoutFoundationITPi = xl.GetData(rowCounter, foundationColumnNumber + 1),
                            .PiApproach = xl.GetData(rowCounter, foundationColumnNumber + 2),
                            .Practice = practice
                        }
                        If jscoreList Is Nothing Then jscoreList = New Dictionary(Of String, JScore)
                        If jscoreList.ContainsKey(jscr.EmpID) Then
                            Dim existingScore As JScore = jscoreList(jscr.EmpID)
                            If existingScore.Practice = "MF" Then
                                jscoreList(jscr.EmpID) = jscr
                            End If
                        Else
                            jscoreList.Add(jscr.EmpID, jscr)
                        End If
                    Next
                End Using
            End If
        Next
        If jscoreList IsNot Nothing AndAlso jscoreList.Count > 0 Then
            Dim jscoreData(jscoreList.Count, 4) As Object
            Dim rowCounter As Integer = 0
            jscoreData(rowCounter, 0) = "EmpID"
            jscoreData(rowCounter, 1) = "With Foundation I T Pi"
            jscoreData(rowCounter, 2) = "Without Foundation I T Pi"
            jscoreData(rowCounter, 3) = "Pi Approach"
            jscoreData(rowCounter, 4) = "Practice"
            rowCounter += 1
            For Each jscore In jscoreList.Values
                jscoreData(rowCounter, 0) = jscore.EmpID
                jscoreData(rowCounter, 1) = jscore.WithFoundationITPi
                jscoreData(rowCounter, 2) = jscore.WithoutFoundationITPi
                jscoreData(rowCounter, 3) = jscore.PiApproach
                jscoreData(rowCounter, 4) = jscore.Practice
                rowCounter += 1
            Next

            OnHeartbeat(String.Format("Opening {0}", _employeeFile))
            Using xl As New ExcelHelper(_employeeFile, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
                AddHandler xl.Heartbeat, AddressOf OnHeartbeat
                AddHandler xl.WaitingFor, AddressOf OnWaitingFor

                OnHeartbeat("Writing JScore data")
                xl.SetActiveSheet("JScore")
                Dim range As String = xl.GetNamedRange(1, jscoreData.GetLength(0) - 1, 1, jscoreData.GetLength(1) - 1)
                xl.WriteArrayToExcel(jscoreData, range)
            End Using
        End If

        OnHeartbeat(String.Format("Opening {0}", _employeeFile))
        Using xl As New ExcelHelper(_employeeFile, ExcelHelper.ExcelOpenStatus.OpenExistingForReadWrite, ExcelHelper.ExcelSaveType.XLS_XLSX, _cts)
            AddHandler xl.Heartbeat, AddressOf OnHeartbeat
            AddHandler xl.WaitingFor, AddressOf OnWaitingFor

            Dim allSheets As List(Of String) = xl.GetExcelSheetsName()
            If allSheets IsNot Nothing AndAlso allSheets.Count > 0 Then
                Dim dataSheet As String = Nothing
                For Each runningSheet In allSheets
                    If runningSheet.Contains("BFSI") Then
                        dataSheet = runningSheet
                        Exit For
                    End If
                Next
                If dataSheet IsNot Nothing Then
                    OnHeartbeat("Modifying master data")
                    xl.SetActiveSheet(dataSheet)
                    Dim dataSheetRange As String = xl.GetNamedRange(1, xl.GetLastRow(1) - 1, 1, xl.GetLastCol(1) - 1)

                    Dim findRange As String = xl.GetNamedRange(1, 1, 1, xl.GetLastCol(1) - 1)
                    Dim applicableColumn As Integer = xl.FindAll("Applicable for WFT", findRange, True).FirstOrDefault.Value
                    Dim statusColumn As Integer = xl.FindAll("Status With Foundation", findRange, True).FirstOrDefault.Value
                    xl.FilterData(statusColumn, "=", "#N/A")
                    xl.FilterData(applicableColumn, "=", "Yes", True)

                    Dim filterdRowList As List(Of Integer) = xl.GetDataRows()
                    If filterdRowList IsNot Nothing AndAlso filterdRowList.Count > 0 Then
                        Dim exclusionColumn As Integer = xl.FindAll("Exclusion Reason", findRange, True).FirstOrDefault.Value
                        For Each row In filterdRowList
                            If row <> 1 Then
                                xl.SetData(row, applicableColumn, "No")
                                xl.SetData(row, exclusionColumn, "No JScore available")
                            End If
                        Next
                    End If
                    xl.UnFilterSheet(dataSheet)

                    xl.CreateNewSheet("Account View")
                    xl.CreateNewSheet("Vertical View")
                    xl.CreateNewSheet("Practice View")
                    xl.CreateNewSheet("Progress Check")
                    Dim orderList As List(Of String) = New List(Of String) From {"Foundation Pending", "Foundation Complete", "I", "T", "Pi"}

                    OnHeartbeat("Creating pivot table 'Vertical View'")
                    xl.CreatPivotTable(dataSheet, dataSheetRange, "Vertical View", "A5",
                                       New List(Of String) From {"Status With Foundation"},
                                       New List(Of String) From {"Vertical"},
                                       New Dictionary(Of String, ExcelHelper.XLFunction) From {{"TOTAL_DAYS", ExcelHelper.XLFunction.Sum}},
                                       New Dictionary(Of String, String) From {{"Applicable for WFT", "Yes"}})
                    xl.ReorderPivotTable("Vertical View", "Status With Foundation", orderList)
                    'Dim verticalOrderList As List(Of String) = New List(Of String) From {"BFS AMERICAS", "BFS EMEA", "SECURITIES", "INSURANCE", "FS-CITI", "SUB", "BANKING PRODUCTS", "BFS- APAC", "BFSI-GRP", "EMERGING MARKETS", "BP ACCOUNT"}
                    Dim verticalOrderList As List(Of String) = New List(Of String) From {"BFS AMERICAS", "BFS EMEA", "SECURITIES", "INSURANCE", "FS-CITI", "BANKING PRODUCTS", "BFS- APAC", "BFSI-GRP", "EMERGING MARKETS", "BP ACCOUNT"}
                    xl.ReorderPivotTable("Vertical View", "Vertical", verticalOrderList)

                    OnHeartbeat("Creating pivot table 'Account View'")
                    xl.CreatPivotTable(dataSheet, dataSheetRange, "Account View", "A5",
                                       New List(Of String) From {"Status With Foundation"},
                                       New List(Of String) From {"GROUP_CUSTOMER_NAME"},
                                       New Dictionary(Of String, ExcelHelper.XLFunction) From {{"TOTAL_DAYS", ExcelHelper.XLFunction.Sum}},
                                       New Dictionary(Of String, String) From {{"Applicable for WFT", "Yes"}})
                    xl.ReorderPivotTable("Account View", "Status With Foundation", orderList)

                    OnHeartbeat("Creating pivot table 'Practice View'")
                    xl.CreatPivotTable(dataSheet, dataSheetRange, "Practice View", "A5",
                                       New List(Of String) From {"Status With Foundation"},
                                       New List(Of String) From {"Practice2"},
                                       New Dictionary(Of String, ExcelHelper.XLFunction) From {{"TOTAL_DAYS", ExcelHelper.XLFunction.Sum}},
                                       New Dictionary(Of String, String) From {{"Applicable for WFT", "Yes"}})
                    xl.ReorderPivotTable("Practice View", "Status With Foundation", orderList)

                    OnHeartbeat("Creating pivot table 'Progress Check'")
                    xl.CreatPivotTable(dataSheet, dataSheetRange, "Progress Check", "A5",
                                      New List(Of String) From {"Last Month Level"},
                                      New List(Of String) From {"Status With Foundation"},
                                      New Dictionary(Of String, ExcelHelper.XLFunction) From {{"TOTAL_DAYS", ExcelHelper.XLFunction.Sum}},
                                      New Dictionary(Of String, String) From {{"Applicable for WFT", "Yes"}})
                    xl.ReorderPivotTable("Progress Check", "Last Month Level", orderList)
                    xl.ReorderPivotTable("Progress Check", "Status With Foundation", orderList)

                    xl.SetActiveSheet("Progress Check")
                    xl.SetData(15, 1, "FS")
                    xl.SetData(15, 2, "# of resources whose levels changed")
                    xl.SetData(15, 4, "% of resources whose levels changed")
                    xl.SetData(16, 2, "Up")
                    xl.SetData(16, 3, "Down")
                    xl.SetData(16, 4, "Up")
                    xl.SetData(16, 5, "Down")

                    xl.SetData(17, 1, "Foundation Complete")
                    xl.SetCellFormula(17, 2, "=B8+G8")
                    xl.SetCellFormula(17, 3, "=C7")
                    xl.SetCellFormula(17, 4, "=B17/$B$22")
                    xl.SetCellFormula(17, 5, "=C17/$B$22")

                    xl.SetData(18, 1, "I")
                    xl.SetCellFormula(18, 2, "=SUM(B9:C9)+G9")
                    xl.SetCellFormula(18, 3, "=SUM(D7:D8)")
                    xl.SetCellFormula(18, 4, "=B18/$B$22")
                    xl.SetCellFormula(18, 5, "=C18/$B$22")

                    xl.SetData(19, 1, "T")
                    xl.SetCellFormula(19, 2, "=SUM(B10:D10)+G10")
                    xl.SetCellFormula(19, 3, "=SUM(E7:E9)")
                    xl.SetCellFormula(19, 4, "=B19/$B$22")
                    xl.SetCellFormula(19, 5, "=C19/$B$22")

                    xl.SetData(20, 1, "Pi")
                    xl.SetCellFormula(20, 2, "=SUM(B11:E11)+G11")
                    xl.SetCellFormula(20, 3, "=SUM(F7:F10)")
                    xl.SetCellFormula(20, 4, "=B20/$B$22")
                    xl.SetCellFormula(20, 5, "=C20/$B$22")

                    xl.SetData(21, 1, "Total Changes")
                    xl.SetCellFormula(21, 2, "=SUM(B17:B20)")
                    xl.SetCellFormula(21, 3, "=SUM(C17:C20)")
                    xl.SetCellFormula(21, 4, "=B21/$B$22")
                    xl.SetCellFormula(21, 5, "=C22/$B$22")

                    xl.SetData(22, 1, "Total Headcount")
                    xl.SetCellFormula(22, 2, "=H12")

                    xl.SetActiveSheet(dataSheet)
                    Dim reportedDate As Date = Convert.ToDateTime(xl.GetData(2, 1))
                    'Dim currentMonth As String = dataSheet.Substring(dataSheet.Count - 3, 3).Trim
                    Dim currentMonth As String = reportedDate.ToString("MMM")
                    If _quarterlyProgress IsNot Nothing AndAlso _quarterlyProgress.Count > 0 AndAlso _quarterlyProgress.ContainsKey(currentMonth.ToUpper) Then
                        Dim currentQuarter As String = _quarterlyProgress(currentMonth.ToUpper)
                        Dim fy As String = Nothing
                        If currentQuarter = "Q4" Then
                            fy = Now.Year.ToString.Substring(2, 2)
                        Else
                            fy = (Now.Year + 1).ToString.Substring(2, 2)
                        End If

                        xl.SetActiveSheet("Vertical View")
                        Dim verticalViewRange As String = "$B$7:$G$16"
                        Dim verticalViewData As Object(,) = xl.GetExcelInMemory(verticalViewRange)
                        xl.SetActiveSheet("Progress")
                        Dim previousMonthRange As String = "$C$5:$I$19"
                        Dim currentMonthRange As String = "$J$5:$P$19"
                        Dim currentMonthDataDumpRange As String = "$J$5:$O$14"

                        xl.SetData(3, 3, xl.GetData(3, 10), ExcelHelper.XLAlign.Center)
                        xl.SetData(3, 10, String.Format("{0} FY {1}", currentMonth, fy), ExcelHelper.XLAlign.Center)
                        Dim previousMonthData As Object(,) = xl.GetExcelInMemory(currentMonthRange)
                        xl.WriteArrayToExcel(previousMonthData, previousMonthRange)
                        xl.WriteArrayToExcel(verticalViewData, currentMonthDataDumpRange)

                        xl.SaveExcel()
                        Dim currentMonthData As Object(,) = xl.GetExcelInMemory(currentMonthDataDumpRange)

                        Dim previousQuarter As String = xl.GetData(23, 10).ToString.Substring(0, 2).Trim
                        Dim previousQuarterRange As String = "$C$24:$I$38"
                        Dim currentQuarterRange As String = "$J$24:$P$38"
                        Dim currentQuarterDataDumpRange As String = "$J$24:$O$33"
                        If previousQuarter.ToUpper <> currentQuarter.ToUpper Then
                            xl.SetData(23, 3, xl.GetData(23, 10), ExcelHelper.XLAlign.Center)
                            xl.SetData(23, 10, String.Format("{0} FY {1}", currentQuarter, fy), ExcelHelper.XLAlign.Center)
                            Dim previousQuarterData As Object(,) = xl.GetExcelInMemory(currentQuarterRange)
                            xl.WriteArrayToExcel(previousQuarterData, previousQuarterRange)
                        End If
                        xl.WriteArrayToExcel(currentMonthData, currentQuarterDataDumpRange)

                        Dim previousYear As String = xl.GetData(43, 10).ToString.Substring(7, 2).Trim
                        Dim previousYearRange As String = "$C$43:$I$57"
                        Dim currentYearRange As String = "$J$43:$P$57"
                        Dim currentYearDataDumpRange As String = "$J$43:$O$52"
                        If previousYear.ToUpper <> fy.ToUpper Then
                            xl.SetData(43, 3, xl.GetData(43, 10), ExcelHelper.XLAlign.Center)
                            Dim previousYearData As Object(,) = xl.GetExcelInMemory(currentYearRange)
                            xl.WriteArrayToExcel(previousYearData, previousYearRange)
                        End If
                        xl.SetData(43, 10, String.Format("{0} FY {1}", currentMonth, fy), ExcelHelper.XLAlign.Center)
                        xl.WriteArrayToExcel(currentMonthData, currentYearDataDumpRange)

                        Dim currentInitiationRange As String = "$J$62:$O$71"
                        xl.SetData(63, 10, String.Format("{0} FY {1}", currentMonth, fy), ExcelHelper.XLAlign.Center)
                        xl.WriteArrayToExcel(currentMonthData, currentInitiationRange)
                    End If
                End If
            End If
        End Using
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
