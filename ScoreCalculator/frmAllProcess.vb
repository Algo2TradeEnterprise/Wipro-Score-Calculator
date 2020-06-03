Imports System.IO
Imports System.Threading

Public Class frmAllProcess

#Region "Common Delegates"
    Delegate Sub SetObjectEnableDisable_Delegate(ByVal [obj] As Object, ByVal [value] As Boolean)
    Public Sub SetObjectEnableDisable_ThreadSafe(ByVal [obj] As Object, ByVal [value] As Boolean)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [obj].InvokeRequired Then
            Dim MyDelegate As New SetObjectEnableDisable_Delegate(AddressOf SetObjectEnableDisable_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[obj], [value]})
        Else
            [obj].Enabled = [value]
        End If
    End Sub

    Delegate Sub SetObjectVisible_Delegate(ByVal [obj] As Object, ByVal [value] As Boolean)
    Public Sub SetObjectVisible_ThreadSafe(ByVal [obj] As Object, ByVal [value] As Boolean)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [obj].InvokeRequired Then
            Dim MyDelegate As New SetObjectVisible_Delegate(AddressOf SetObjectVisible_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[obj], [value]})
        Else
            [obj].Visible = [value]
        End If
    End Sub

    Delegate Sub SetLabelText_Delegate(ByVal [label] As Label, ByVal [text] As String)
    Public Sub SetLabelText_ThreadSafe(ByVal [label] As Label, ByVal [text] As String)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [label].InvokeRequired Then
            Dim MyDelegate As New SetLabelText_Delegate(AddressOf SetLabelText_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[label], [text]})
        Else
            [label].Text = [text]
        End If
    End Sub

    Delegate Function GetLabelText_Delegate(ByVal [label] As Label) As String
    Public Function GetLabelText_ThreadSafe(ByVal [label] As Label) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [label].InvokeRequired Then
            Dim MyDelegate As New GetLabelText_Delegate(AddressOf GetLabelText_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[label]})
        Else
            Return [label].Text
        End If
    End Function

    Delegate Sub SetLabelTag_Delegate(ByVal [label] As Label, ByVal [tag] As String)
    Public Sub SetLabelTag_ThreadSafe(ByVal [label] As Label, ByVal [tag] As String)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [label].InvokeRequired Then
            Dim MyDelegate As New SetLabelTag_Delegate(AddressOf SetLabelTag_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[label], [tag]})
        Else
            [label].Tag = [tag]
        End If
    End Sub

    Delegate Function GetLabelTag_Delegate(ByVal [label] As Label) As String
    Public Function GetLabelTag_ThreadSafe(ByVal [label] As Label) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [label].InvokeRequired Then
            Dim MyDelegate As New GetLabelTag_Delegate(AddressOf GetLabelTag_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[label]})
        Else
            Return [label].Tag
        End If
    End Function
    Delegate Sub SetToolStripLabel_Delegate(ByVal [toolStrip] As StatusStrip, ByVal [label] As ToolStripStatusLabel, ByVal [text] As String)
    Public Sub SetToolStripLabel_ThreadSafe(ByVal [toolStrip] As StatusStrip, ByVal [label] As ToolStripStatusLabel, ByVal [text] As String)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [toolStrip].InvokeRequired Then
            Dim MyDelegate As New SetToolStripLabel_Delegate(AddressOf SetToolStripLabel_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[toolStrip], [label], [text]})
        Else
            [label].Text = [text]
        End If
    End Sub

    Delegate Function GetToolStripLabel_Delegate(ByVal [toolStrip] As StatusStrip, ByVal [label] As ToolStripLabel) As String
    Public Function GetToolStripLabel_ThreadSafe(ByVal [toolStrip] As StatusStrip, ByVal [label] As ToolStripLabel) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [toolStrip].InvokeRequired Then
            Dim MyDelegate As New GetToolStripLabel_Delegate(AddressOf GetToolStripLabel_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[toolStrip], [label]})
        Else
            Return [label].Text
        End If
    End Function

    Delegate Function GetDateTimePickerValue_Delegate(ByVal [dateTimePicker] As DateTimePicker) As Date
    Public Function GetDateTimePickerValue_ThreadSafe(ByVal [dateTimePicker] As DateTimePicker) As Date
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [dateTimePicker].InvokeRequired Then
            Dim MyDelegate As New GetDateTimePickerValue_Delegate(AddressOf GetDateTimePickerValue_ThreadSafe)
            Return Me.Invoke(MyDelegate, New DateTimePicker() {[dateTimePicker]})
        Else
            Return [dateTimePicker].Value
        End If
    End Function

    Delegate Function GetNumericUpDownValue_Delegate(ByVal [numericUpDown] As NumericUpDown) As Integer
    Public Function GetNumericUpDownValue_ThreadSafe(ByVal [numericUpDown] As NumericUpDown) As Integer
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [numericUpDown].InvokeRequired Then
            Dim MyDelegate As New GetNumericUpDownValue_Delegate(AddressOf GetNumericUpDownValue_ThreadSafe)
            Return Me.Invoke(MyDelegate, New NumericUpDown() {[numericUpDown]})
        Else
            Return [numericUpDown].Value
        End If
    End Function

    Delegate Function GetComboBoxIndex_Delegate(ByVal [combobox] As ComboBox) As Integer
    Public Function GetComboBoxIndex_ThreadSafe(ByVal [combobox] As ComboBox) As Integer
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [combobox].InvokeRequired Then
            Dim MyDelegate As New GetComboBoxIndex_Delegate(AddressOf GetComboBoxIndex_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[combobox]})
        Else
            Return [combobox].SelectedIndex
        End If
    End Function

    Delegate Function GetComboBoxItem_Delegate(ByVal [ComboBox] As ComboBox) As String
    Public Function GetComboBoxItem_ThreadSafe(ByVal [ComboBox] As ComboBox) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [ComboBox].InvokeRequired Then
            Dim MyDelegate As New GetComboBoxItem_Delegate(AddressOf GetComboBoxItem_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[ComboBox]})
        Else
            Return [ComboBox].SelectedItem.ToString
        End If
    End Function

    Delegate Function GetTextBoxText_Delegate(ByVal [textBox] As TextBox) As String
    Public Function GetTextBoxText_ThreadSafe(ByVal [textBox] As TextBox) As String
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [textBox].InvokeRequired Then
            Dim MyDelegate As New GetTextBoxText_Delegate(AddressOf GetTextBoxText_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[textBox]})
        Else
            Return [textBox].Text
        End If
    End Function

    Delegate Function GetCheckBoxChecked_Delegate(ByVal [checkBox] As CheckBox) As Boolean
    Public Function GetCheckBoxChecked_ThreadSafe(ByVal [checkBox] As CheckBox) As Boolean
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [checkBox].InvokeRequired Then
            Dim MyDelegate As New GetCheckBoxChecked_Delegate(AddressOf GetCheckBoxChecked_ThreadSafe)
            Return Me.Invoke(MyDelegate, New Object() {[checkBox]})
        Else
            Return [checkBox].Checked
        End If
    End Function

    Delegate Sub SetDatagridBindDatatable_Delegate(ByVal [datagrid] As DataGridView, ByVal [table] As DataTable)
    Public Sub SetDatagridBindDatatable_ThreadSafe(ByVal [datagrid] As DataGridView, ByVal [table] As DataTable)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [datagrid].InvokeRequired Then
            Dim MyDelegate As New SetDatagridBindDatatable_Delegate(AddressOf SetDatagridBindDatatable_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {[datagrid], [table]})
        Else
            [datagrid].DataSource = [table]
            [datagrid].Refresh()
        End If
    End Sub

    Delegate Sub SetListAddItem_Delegate(ByVal [lst] As ListBox, ByVal [value] As Object)
    Public Sub SetListAddItem_ThreadSafe(ByVal [lst] As ListBox, ByVal [value] As Object)
        ' InvokeRequired required compares the thread ID of the calling thread to the thread ID of the creating thread.  
        ' If these threads are different, it returns true.  
        If [lst].InvokeRequired Then
            Dim MyDelegate As New SetListAddItem_Delegate(AddressOf SetListAddItem_ThreadSafe)
            Me.Invoke(MyDelegate, New Object() {lst, [value]})
        Else
            [lst].Items.Insert(0, [value])
        End If
    End Sub
#End Region

#Region "Event Handlers"
    Private Sub OnHeartbeatMain(message As String)
        SetLabelText_ThreadSafe(lblMainProgress, message)
    End Sub
    Private Sub OnHeartbeat(message As String)
        SetLabelText_ThreadSafe(lblProgress, message)
    End Sub
    Private Sub OnHeartbeatError(message As String)
        SetListAddItem_ThreadSafe(lstError, message)
    End Sub
    Private Sub OnDocumentDownloadComplete()
        'OnHeartbeat("Document download compelete")
    End Sub
    Private Sub OnDocumentRetryStatus(currentTry As Integer, totalTries As Integer)
        OnHeartbeat(String.Format("Try #{0}/{1}: Connecting...", currentTry, totalTries))
    End Sub
    Public Sub OnWaitingFor(ByVal elapsedSecs As Integer, ByVal totalSecs As Integer, ByVal msg As String)
        OnHeartbeat(String.Format("{0}, waiting {1}/{2} secs", msg, elapsedSecs, totalSecs))
    End Sub
#End Region

    Private ReadOnly _preProcessFolder As String = Path.Combine(My.Application.Info.DirectoryPath, "Excel Test", "Pre Process")
    Private ReadOnly _preProcessScoreModifierFolder As String = Path.Combine(My.Application.Info.DirectoryPath, "Excel Test", "Pre Process", "Score Modifier")
    Private ReadOnly _inProcessFolder As String = Path.Combine(My.Application.Info.DirectoryPath, "Excel Test", "In Process")
    Private ReadOnly _postProcessFolder As String = Path.Combine(My.Application.Info.DirectoryPath, "Excel Test", "Post Process")

    Private _addGraceMark As Boolean = False
    Private _canceller As CancellationTokenSource
    Private Sub frmPreProcess_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtFolderpath.Text = My.Settings.FolderPath
        chkbZeroScoreReplacement.Checked = My.Settings.ZeroScoreReplace
        chkbMaxScoreReplacement.Checked = My.Settings.MaxScoreReplace
        chkbFndtnCmplt.Checked = My.Settings.FoundationPendingToComplete
        chkbITPi.Checked = My.Settings.FoundationCompleteToITPi
        chkbFndtnCmplt_CheckedChanged(sender, e)
        chkbITPi_CheckedChanged(sender, e)
        txtFndtnCmplt.Text = My.Settings.FoundationPendingToCompleteMinScore
        txtITPi.Text = My.Settings.FoundationCompleteToITPiMinScore
        chkbAutoProcess.Checked = My.Settings.AutoProcessAllStep

        SetObjectEnableDisable_ThreadSafe(btnStop, False)
    End Sub

    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        If folderBrowse.ShowDialog() = DialogResult.OK Then
            txtFolderpath.Text = folderBrowse.SelectedPath
        End If
    End Sub

    Private Sub btnStop_Click(sender As Object, e As EventArgs) Handles btnStop.Click
        _canceller.Cancel()
    End Sub

    Private Async Sub btnStartWithScoreAdjustment_Click(sender As Object, e As EventArgs) Handles btnStartWithScoreAdjustment.Click
        _addGraceMark = True
        My.Settings.FolderPath = txtFolderpath.Text
        My.Settings.AutoProcessAllStep = chkbAutoProcess.Checked
        My.Settings.Save()
        Await Task.Run(AddressOf StartProcessingAsync).ConfigureAwait(False)
    End Sub

    Private Async Sub btnStartWithoutScoreAdjustment_Click(sender As Object, e As EventArgs) Handles btnStartWithoutScoreAdjustment.Click
        _addGraceMark = False
        My.Settings.FolderPath = txtFolderpath.Text
        My.Settings.AutoProcessAllStep = chkbAutoProcess.Checked
        My.Settings.Save()
        Await Task.Run(AddressOf StartProcessingAsync).ConfigureAwait(False)
    End Sub

    Private Async Function StartProcessingAsync() As Task
        SetObjectEnableDisable_ThreadSafe(grpFolderBrowse, False)
        SetObjectEnableDisable_ThreadSafe(btnStartWithScoreAdjustment, False)
        SetObjectEnableDisable_ThreadSafe(btnStartWithoutScoreAdjustment, False)
        SetObjectEnableDisable_ThreadSafe(btnStop, True)
        Try
            _canceller = New CancellationTokenSource

            'Pre Process
            StartFileDistribution()
            Await StartScoreModifyProcessingAsync().ConfigureAwait(False)
            Await StartMaxScoreReplacementProcessingAsync(_addGraceMark).ConfigureAwait(False)

            If GetCheckBoxChecked_ThreadSafe(chkbAutoProcess) Then
                'In Process
                Await StartInProcessing().ConfigureAwait(False)
                'Post Process
                Await StartPostProcessing().ConfigureAwait(False)
            Else
                If MessageBox.Show("Pre process complete. Proceed to next step?", "Score Calculator", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                    'In Process
                    Await StartInProcessing().ConfigureAwait(False)
                    If MessageBox.Show("In process complete. Proceed to next step?", "Score Calculator", MessageBoxButtons.YesNo) = DialogResult.Yes Then
                        'Post Process
                        Await StartPostProcessing().ConfigureAwait(False)
                    End If
                End If
            End If
        Catch cex As OperationCanceledException
            MessageBox.Show(cex.Message, "Score Calculator", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Catch ex As Exception
            MessageBox.Show(ex.ToString, "Score Calculator", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            SetObjectEnableDisable_ThreadSafe(grpFolderBrowse, True)
            SetObjectEnableDisable_ThreadSafe(btnStartWithScoreAdjustment, True)
            SetObjectEnableDisable_ThreadSafe(btnStartWithoutScoreAdjustment, True)
            SetObjectEnableDisable_ThreadSafe(btnStop, False)
            OnHeartbeatMain("Process complete")
            OnHeartbeat("")
        End Try
    End Function

    Private Sub StartFileDistribution()
        OnHeartbeatMain("Distributing files to there required folders")
        Dim inputFolder As String = txtFolderpath.Text
        For Each runningFile In Directory.GetFiles(inputFolder)
            _canceller.Token.ThrowIfCancellationRequested()
            If runningFile.ToUpper.Contains("MAPPING") Then
                Dim preProcessMappingFile As String = Path.Combine(_preProcessFolder, Path.GetFileName(runningFile))
                If File.Exists(preProcessMappingFile) Then File.Delete(preProcessMappingFile)
                File.Copy(runningFile, preProcessMappingFile)

                Dim inProcessMappingFile As String = Path.Combine(_inProcessFolder, Path.GetFileName(runningFile))
                If File.Exists(inProcessMappingFile) Then File.Delete(inProcessMappingFile)
                File.Copy(runningFile, inProcessMappingFile)
            ElseIf runningFile.ToUpper.Contains("ASG") Then
                Dim inProcessFile As String = Path.Combine(_inProcessFolder, Path.GetFileName(runningFile))
                If File.Exists(inProcessFile) Then File.Delete(inProcessFile)
                File.Copy(runningFile, inProcessFile)
            ElseIf runningFile.ToUpper.Contains("BFSI") Then
                Dim preProcessFile As String = Path.Combine(_preProcessFolder, Path.GetFileName(runningFile))
                If File.Exists(preProcessFile) Then File.Delete(preProcessFile)
                File.Copy(runningFile, preProcessFile)
            End If
        Next

        Dim previousMonth As String = Nothing
        For Each runningFile In Directory.GetFiles(inputFolder)
            _canceller.Token.ThrowIfCancellationRequested()
            If runningFile.ToUpper.Contains("OUTPUT") Then
                Dim preProcessScrMdfrFile As String = Path.Combine(_preProcessScoreModifierFolder, Path.GetFileName(runningFile))
                If File.Exists(preProcessScrMdfrFile) Then File.Delete(preProcessScrMdfrFile)
                File.Copy(runningFile, preProcessScrMdfrFile)

                previousMonth = Path.GetFileName(preProcessScrMdfrFile).Substring(Path.GetFileName(preProcessScrMdfrFile).Count - 13)
            End If
        Next

        Dim mdfyPndngToCmplt As Boolean = GetCheckBoxChecked_ThreadSafe(chkbFndtnCmplt)
        Dim mdfyCmpltToITPi As Boolean = GetCheckBoxChecked_ThreadSafe(chkbITPi)
        If previousMonth IsNot Nothing Then
            For Each runningFile In Directory.GetFiles(inputFolder)
                _canceller.Token.ThrowIfCancellationRequested()
                If runningFile.ToUpper.Contains("ANALYSIS") Then
                    If runningFile.ToUpper.Contains(previousMonth.ToUpper) Then
                        Dim preProcessFile As String = Path.Combine(_preProcessFolder, Path.GetFileName(runningFile))
                        If File.Exists(preProcessFile) Then File.Delete(preProcessFile)
                        File.Copy(runningFile, preProcessFile)
                    Else
                        If mdfyPndngToCmplt OrElse mdfyCmpltToITPi Then
                            Dim preProcessScrMdfrFile As String = Path.Combine(_preProcessScoreModifierFolder, Path.GetFileName(runningFile))
                            If File.Exists(preProcessScrMdfrFile) Then File.Delete(preProcessScrMdfrFile)
                            File.Copy(runningFile, preProcessScrMdfrFile)
                        Else
                            Dim preProcessFile As String = Path.Combine(_preProcessFolder, Path.GetFileName(runningFile))
                            If File.Exists(preProcessFile) Then File.Delete(preProcessFile)
                            File.Copy(runningFile, preProcessFile)
                        End If
                    End If
                End If
            Next
        End If
    End Sub

    Private Async Function StartScoreModifyProcessingAsync() As Task
        OnHeartbeatMain("Checking Foundation Complete and I T Pi upgrade from Foundation Pending and Foundation Complete")
        Dim mdfyPndngToCmplt As Boolean = GetCheckBoxChecked_ThreadSafe(chkbFndtnCmplt)
        Dim mdfyPndngToCmpltMinScore As Decimal = GetTextBoxText_ThreadSafe(txtFndtnCmplt)
        Dim mdfyCmpltToITPi As Boolean = GetCheckBoxChecked_ThreadSafe(chkbITPi)
        Dim mdfyCmpltToITPiMinScore As Decimal = GetTextBoxText_ThreadSafe(txtITPi)

        My.Settings.FoundationPendingToComplete = mdfyPndngToCmplt
        My.Settings.FoundationCompleteToITPi = mdfyCmpltToITPi
        My.Settings.FoundationPendingToCompleteMinScore = mdfyPndngToCmpltMinScore
        My.Settings.FoundationCompleteToITPiMinScore = mdfyCmpltToITPiMinScore
        My.Settings.Save()

        If mdfyPndngToCmplt OrElse mdfyCmpltToITPi Then
            Dim directoryPath As String = _preProcessScoreModifierFolder
            Dim empFile As String = Nothing

            Using scrMdfr As New ScoreModifier(_canceller, directoryPath, mdfyPndngToCmplt, mdfyCmpltToITPi, mdfyPndngToCmpltMinScore, mdfyCmpltToITPiMinScore)
                AddHandler scrMdfr.Heartbeat, AddressOf OnHeartbeat
                AddHandler scrMdfr.HeartbeatError, AddressOf OnHeartbeatError
                AddHandler scrMdfr.WaitingFor, AddressOf OnWaitingFor
                AddHandler scrMdfr.DocumentRetryStatus, AddressOf OnDocumentRetryStatus
                AddHandler scrMdfr.DocumentDownloadComplete, AddressOf OnDocumentDownloadComplete

                Await scrMdfr.ProcessDataAsync().ConfigureAwait(False)
            End Using
        End If
    End Function

    Private Async Function StartMaxScoreReplacementProcessingAsync(ByVal addGraceMark As Boolean) As Task
        Dim zeroScoreReplace As Boolean = GetCheckBoxChecked_ThreadSafe(chkbZeroScoreReplacement)
        Dim maxScoreReplace As Boolean = GetCheckBoxChecked_ThreadSafe(chkbMaxScoreReplacement)
        My.Settings.ZeroScoreReplace = zeroScoreReplace
        My.Settings.MaxScoreReplace = maxScoreReplace
        My.Settings.Save()

        If zeroScoreReplace OrElse maxScoreReplace Then
            Dim folderPath As String = Nothing
            Dim mappingFile As String = Nothing
            Dim directoryName As String = Path.Combine(My.Application.Info.DirectoryPath, "Excel Test", "Pre Process")
            folderPath = directoryName
            For Each runningFile In Directory.GetFiles(directoryName)
                If runningFile.ToUpper.Contains("MAPPING") Then
                    mappingFile = runningFile
                End If
            Next
            If mappingFile Is Nothing Then Throw New ApplicationException("Mapping file not exits")

            Using prePrcsHlpr As New PreProcess(_canceller, folderPath, mappingFile, zeroScoreReplace, maxScoreReplace, addGraceMark)
                AddHandler prePrcsHlpr.Heartbeat, AddressOf OnHeartbeat
                AddHandler prePrcsHlpr.HeartbeatError, AddressOf OnHeartbeatError
                AddHandler prePrcsHlpr.WaitingFor, AddressOf OnWaitingFor
                AddHandler prePrcsHlpr.DocumentRetryStatus, AddressOf OnDocumentRetryStatus
                AddHandler prePrcsHlpr.DocumentDownloadComplete, AddressOf OnDocumentDownloadComplete

                OnHeartbeatMain("Copying necessary data and sheets from previous moth BFSI file")
                Await prePrcsHlpr.ProcessEmployeeData().ConfigureAwait(False)
                OnHeartbeatMain("Updating zero score and max score data from previous month score")
                Await prePrcsHlpr.ProcessScoreData().ConfigureAwait(False)
            End Using
        End If
    End Function

    Private Async Function StartInProcessing() As Task
        OnHeartbeatMain("Calculating Foundation and I T Pi scores of all pratice")
        Dim mappingFilename As String = Nothing
        Dim empFilename As String = Nothing
        Dim asgFilename As String = Nothing
        Dim directoryName As String = Path.Combine(My.Application.Info.DirectoryPath, "Excel Test", "In Process")
        For Each runningFile In Directory.GetFiles(directoryName)
            If runningFile.ToUpper.Contains("MAPPING") Then
                mappingFilename = runningFile
            ElseIf runningFile.ToUpper.Contains("BFSI") Then
                empFilename = runningFile
            ElseIf runningFile.ToUpper.Contains("ASG") Then
                asgFilename = runningFile
            End If
        Next

        If mappingFilename Is Nothing Then Throw New ApplicationException("Mapping file not exits")
        If empFilename Is Nothing Then Throw New ApplicationException("Employee file not exits")
        Using inPrcsHlpr As New InProcess(_canceller, mappingFilename, empFilename, asgFilename)
            AddHandler inPrcsHlpr.Heartbeat, AddressOf OnHeartbeat
            AddHandler inPrcsHlpr.HeartbeatError, AddressOf OnHeartbeatError
            AddHandler inPrcsHlpr.WaitingFor, AddressOf OnWaitingFor
            AddHandler inPrcsHlpr.DocumentRetryStatus, AddressOf OnDocumentRetryStatus
            AddHandler inPrcsHlpr.DocumentDownloadComplete, AddressOf OnDocumentDownloadComplete

            Await inPrcsHlpr.ProcessData().ConfigureAwait(False)
        End Using
    End Function

    Private Async Function StartPostProcessing() As Task
        OnHeartbeatMain("Merging all data and creating neccesary tables at BFSI file")
        Dim empFilename As String = Nothing
        Dim directoryName As String = Path.Combine(My.Application.Info.DirectoryPath, "Excel Test", "Post Process")
        For Each runningFile In Directory.GetFiles(directoryName)
            If runningFile.ToUpper.Contains("BFSI") Then
                empFilename = runningFile
            End If
        Next
        If empFilename Is Nothing Then Throw New ApplicationException("Employee file not exits")

        Using pstPrcsHlpr As New PostProcess(_canceller, empFilename)
            AddHandler pstPrcsHlpr.Heartbeat, AddressOf OnHeartbeat
            AddHandler pstPrcsHlpr.HeartbeatError, AddressOf OnHeartbeatError
            AddHandler pstPrcsHlpr.WaitingFor, AddressOf OnWaitingFor
            AddHandler pstPrcsHlpr.DocumentRetryStatus, AddressOf OnDocumentRetryStatus
            AddHandler pstPrcsHlpr.DocumentDownloadComplete, AddressOf OnDocumentDownloadComplete

            Await pstPrcsHlpr.ProcessData().ConfigureAwait(False)
        End Using
    End Function

    Private Sub chkbFndtnCmplt_CheckedChanged(sender As Object, e As EventArgs) Handles chkbFndtnCmplt.CheckedChanged
        lblFndtnCmplt.Visible = chkbFndtnCmplt.Checked
        txtFndtnCmplt.Visible = chkbFndtnCmplt.Checked
    End Sub

    Private Sub chkbITPi_CheckedChanged(sender As Object, e As EventArgs) Handles chkbITPi.CheckedChanged
        lblITPi.Visible = chkbITPi.Checked
        txtITPi.Visible = chkbITPi.Checked
    End Sub
End Class