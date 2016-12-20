Imports System.IO
Imports System.Threading

Public Class Form1
    Public Const settingsFilePath = "settings.ini"
#Region "Basic functions"
    Private Function convertBytesToAppropriateScale(bytes As Long, Optional precision As Integer = 2)
        If bytes < 1024 Then
            Return String.Format("{0} bytes", bytes)
        ElseIf bytes < (1024 ^ 2) Then
            Return String.Format("{0} kilobytes", Math.Round(bytes / 1024, precision))
        ElseIf bytes < (1024 ^ 3) Then
            Return String.Format("{0} megabytes", Math.Round(bytes / (1024 ^ 2), precision))
        ElseIf bytes < (1024 ^ 4) Then
            Return String.Format("{0} gigabytes", Math.Round(bytes / (1024 ^ 3), precision))
        ElseIf bytes < (1024 ^ 5) Then
            Return String.Format("{0} terrabytes", Math.Round(bytes / (1024 ^ 4), precision))
        Else
            Return bytes.ToString
        End If
    End Function

    Private Sub addFileToList(filePath As String)
        If (My.Computer.FileSystem.FileExists(filePath)) Then
            Dim currentFile As New FileInfo(filePath)

            Dim supported As Boolean = False
            For Each fileType In supportedFormats.Keys
                If supportedFormats(fileType).Contains(currentFile.Extension.ToLower) Then
                    supported = True
                End If
            Next
            If Not supported Then
                Exit Sub
            End If

            Dim newItem As New ListViewItem
            newItem.Text = currentFile.Name
            newItem.Tag = currentFile.FullName
            newItem.SubItems.Add(convertBytesToAppropriateScale(currentFile.Length))
            newItem.SubItems.Item(1).Tag = currentFile.Length
            ListView1.Items.Add(newItem)
        ElseIf My.Computer.FileSystem.DirectoryExists(filePath) Then
            For Each Filepath1 In My.Computer.FileSystem.GetFiles(filePath)
                addFileToList(Filepath1)
            Next
            For Each Dirpath In My.Computer.FileSystem.GetDirectories(filePath)
                addFileToList(Dirpath)
            Next
        End If
    End Sub

    Private Sub saveSettings()
        Dim settingsDict As New Dictionary(Of String, String)
        settingsDict.Add("processPriority", processPriority)
        settingsDict.Add("copyExif", copyExif)
        settingsDict.Add("optimiseHuffmanTable", optimiseHuffmanTable)
        settingsDict.Add("convertToProgressive", convertToProgressive)
        settingsDict.Add("arithmeticCoding", arithmeticCoding)
        settingsDict.Add("saveLocation", saveLocation)
        SettingsFile.saveSettings(settingsDict, settingsFilePath)
    End Sub
    Private Sub loadSettings()
        Dim settingsDict As Dictionary(Of String, String) = SettingsFile.loadSettings(settingsFilePath)
        
        If settingsDict.ContainsKey("processPriority") Then
            processPriority = Int(settingsDict("processPriority"))
        End If
        If settingsDict.ContainsKey("copyExif") Then
            copyExif = Int(settingsDict("copyExif"))
        End If
        If settingsDict.ContainsKey("optimiseHuffmanTable") Then
            optimiseHuffmanTable = CType(settingsDict("optimiseHuffmanTable"), Boolean)
        End If
        If settingsDict.ContainsKey("convertToProgressive") Then
            convertToProgressive = CType(settingsDict("convertToProgressive"), Boolean)
        End If
        If settingsDict.ContainsKey("arithmeticCoding") Then
            arithmeticCoding = CType(settingsDict("arithmeticCoding"), Boolean)
        End If
        If settingsDict.ContainsKey("saveLocation") Then
            saveLocation = settingsDict("saveLocation")
        End If
    End Sub
#End Region

#Region "Optimise"
    Public supportedFormats As New Dictionary(Of String, String()) From {{"JPEG", {".jpg", ".jpeg"}}, _
                                                                      {"PNG", {".png"}}}
    Public processPriority = ProcessPriorityClass.Idle
    Public copyExif As Integer = 1
    Public optimiseHuffmanTable As Boolean = True
    Public convertToProgressive As Boolean = True
    Public arithmeticCoding As Boolean = False
    Public saveLocation As String = ""
    Public Const JPEGtranLocation As String = "bins\jpegtran.exe"
    Public Const OptiPngLocation As String = "bins\optipng0.7.6.exe"
    Public Const PngoutLocation As String = "bins\pngout.exe"
    Public Const DefloptLocation As String = "bins\DeflOpt.exe"

    Private Sub optimiseImages()
        Button1.Enabled = False
        For i As Integer = 0 To ListView1.Items.Count - 1
            Dim item As ListViewItem = ListView1.Items.Item(i)
            Dim currentFile As New FileInfo(item.Tag)

            'Get current path without extention
            Dim outputFilename As String
            If ToolStripMenuItem2.CheckState = CheckState.Checked Then
                outputFilename = currentFile.FullName.Substring(0, currentFile.FullName.Length - currentFile.Extension.Length)
            Else
                outputFilename = saveLocation + currentFile.Name.Substring(0, currentFile.Name.Length - currentFile.Extension.Length)
            End If

            'Detirmine file type and apply appropriate optimisation
            If supportedFormats("JPEG").Contains(currentFile.Extension.ToLower) Then
                Dim jpegOptimWCB As WaitCallback = New WaitCallback(AddressOf optimiseJPEGthread)
                ThreadPool.QueueUserWorkItem(jpegOptimWCB, {currentFile.FullName, outputFilename, i})
            ElseIf supportedFormats("PNG").Contains(currentFile.Extension.ToLower) Then
                Dim pngOptimWCB As WaitCallback = New WaitCallback(AddressOf optimisePNGthread)
                ThreadPool.QueueUserWorkItem(pngOptimWCB, {currentFile.FullName, outputFilename, i})
            End If
        Next
    End Sub

    Private Sub optimiseJPEGthread(parameters As Array)
        optimiseJPEG(parameters(0), parameters(1), parameters(2))

        'Re-enable optimise button if this is the last image that has been optimised
        reenableOptimiseButtonOnJobFinish()
    End Sub
    Private Sub optimiseJPEG(inputFilePath As String, outputFilePath As String, itemIndex As Integer)
        Dim inputFileInfo As New FileInfo(inputFilePath)
        outputFilePath = getNonExistingPath(outputFilePath, inputFileInfo.Extension)

        updateProcessingListViewItem(itemIndex)

        'Exif data
        Dim exifdata
        If copyExif = 0 Then
            exifdata = "-copy none"
        ElseIf copyExif = 1 Then
            exifdata = "-copy comments"
        ElseIf copyExif = 2 Then
            exifdata = "-copy all"
        End If

        'Huffman table
        Dim huffmanTable
        If optimiseHuffmanTable Then
            huffmanTable = "-optimize"
        Else
            huffmanTable = ""
        End If

        'progressive
        Dim progressive
        If convertToProgressive Then
            progressive = "-progressive"
        Else
            progressive = ""
        End If

        Dim arithmetic
        If arithmeticCoding Then
            arithmetic = "-arithmetic"
        Else
            arithmetic = ""
        End If

        Dim JPEGtranProcess As New Process

        With JPEGtranProcess.StartInfo
            .FileName = JPEGtranLocation
            .Arguments = String.Format("-optimize -progressive {2} {3} {4} {5} -outfile {0} {1}", _
                                      """" + outputFilePath + """", """" + inputFilePath + """", exifdata, _
                                      huffmanTable, progressive, arithmetic)
            .UseShellExecute = False
            .CreateNoWindow = True
        End With

        JPEGtranProcess.Start()

        JPEGtranProcess.PriorityClass = processPriority

        While Not JPEGtranProcess.HasExited
            Thread.Sleep(100)
        End While

        Try
            If ToolStripMenuItem2.CheckState = CheckState.Checked Then
                If My.Computer.FileSystem.FileExists(outputFilePath) Then
                    My.Computer.FileSystem.DeleteFile(inputFilePath)
                    My.Computer.FileSystem.RenameFile(outputFilePath, inputFileInfo.Name)
                End If
                updateFinishedListViewItem(itemIndex, inputFilePath)
            Else
                updateFinishedListViewItem(itemIndex, outputFilePath)
            End If

        Catch ex As Exception

        End Try

        Debug.Print(inputFilePath + " Ended")
    End Sub

    Private Sub optimisePNGthread(parameters As Array)
        optimisePNG(parameters(0), parameters(1), parameters(2))

        reenableOptimiseButtonOnJobFinish()
    End Sub

    Private Sub optimisePNG(inputFilePath As String, outputFilePath As String, itemIndex As Integer)
        Dim inputFileInfo As New FileInfo(inputFilePath)
        outputFilePath = getNonExistingPath(outputFilePath, inputFileInfo.Extension)

        updateProcessingListViewItem(itemIndex)

        optiPNGoptimise(inputFileInfo.FullName, outputFilePath)

        pngOUToptimise(outputFilePath, outputFilePath)

        deflOptOptimise(outputFilePath)

        Try
            If ToolStripMenuItem2.CheckState = CheckState.Checked Then
                If My.Computer.FileSystem.FileExists(outputFilePath) Then
                    My.Computer.FileSystem.DeleteFile(inputFilePath)
                    My.Computer.FileSystem.RenameFile(outputFilePath, inputFileInfo.Name)
                End If
                updateFinishedListViewItem(itemIndex, inputFilePath)
            Else
                updateFinishedListViewItem(itemIndex, outputFilePath)
            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub optiPNGoptimise(inputFilePath As String, outputFilePath As String, Optional preserve As Boolean = True, _
                                Optional optimiserLevel As Integer = 7)
        Dim optiPNGprocess As New Process

        Dim arguments As String = ""

        If preserve Then
            arguments += "-preserve "
        End If

        arguments += String.Format("-o{0} ", optimiserLevel)

        arguments += String.Format("-out ""{0}"" ", outputFilePath)

        arguments += """" + inputFilePath + """"

        With optiPNGprocess.StartInfo
            .FileName = OptiPngLocation
            .Arguments = arguments
            .UseShellExecute = False
            .CreateNoWindow = True
        End With

        optiPNGprocess.Start()
        optiPNGprocess.PriorityClass = processPriority

        While Not optiPNGprocess.HasExited
            Thread.Sleep(100)
        End While
    End Sub

    Private Sub pngOUToptimise(inputFilePath As String, outputFilePath As String, Optional overwriteFile As Boolean = True, _
                               Optional bitdepth As Integer = 0, Optional strategy As Integer = 0, _
                               Optional keepChunks As String = "t", Optional mincodes As Integer = 0)
        Dim pngOUTprocess As New Process

        Dim arguments As String = ""

        arguments += String.Format("""{0}""", inputFilePath)
        arguments += String.Format(" ""{0}""", outputFilePath)

        If overwriteFile Then
            arguments += " /y"
        End If

        arguments += String.Format(" /d{0}", bitdepth)

        arguments += String.Format(" /s{0}", strategy)

        If Not keepChunks = String.Empty Then
            arguments += String.Format(" /k{0}", keepChunks)
        End If

        arguments += String.Format(" /mincodes{0}", mincodes)

        With pngOUTprocess.StartInfo
            .FileName = PngoutLocation
            .Arguments = arguments
            .UseShellExecute = False
            .CreateNoWindow = True
        End With

        pngOUTprocess.Start()
        pngOUTprocess.PriorityClass = processPriority

        While Not pngOUTprocess.HasExited
            Thread.Sleep(100)
        End While
    End Sub

    Private Sub deflOptOptimise(inputFilePath As String, Optional preserveDateAndTime As Boolean = True)
        Dim deflOptProcess As New Process

        Dim arguments As String = ""

        If preserveDateAndTime Then
            arguments += "/d "
        End If

        arguments += String.Format("""{0}""", inputFilePath)

        With deflOptProcess.StartInfo
            .FileName = DefloptLocation
            .Arguments = arguments
            .UseShellExecute = False
            .CreateNoWindow = True
        End With

        deflOptProcess.Start()
        deflOptProcess.PriorityClass = processPriority

        While Not deflOptProcess.HasExited
            Thread.Sleep(100)
        End While
    End Sub
    Private Sub reenableOptimiseButtonOnJobFinish()
        Dim worker As Integer = 0
        Dim io As Integer = 0
        ThreadPool.GetAvailableThreads(worker, io)
        If (worker >= Environment.ProcessorCount - 1 And io >= Environment.ProcessorCount - 1) Then
            Me.Invoke(Sub() Button1.Enabled = True)
        End If
    End Sub

    Private Sub updateProcessingListViewItem(itemIndex As Integer)
        Dim currentItem As ListViewItem
        Me.Invoke(Sub() currentItem = ListView1.Items.Item(itemIndex))

        Dim newItem As ListViewItem = currentItem
        'If the listview item is already set to "Done" simply change the value, instead of trying to create new subitems
        Me.Invoke(Sub() ListView1.BeginUpdate())
        If newItem.SubItems.Count > 2 Then
            Me.Invoke(Sub() newItem.SubItems.Item(2).Text = "")
            Me.Invoke(Sub() newItem.SubItems.Item(3).Text = "")
            Me.Invoke(Sub() newItem.SubItems.Item(4).Text = "Processing")
        Else
            Me.Invoke(Sub() newItem.SubItems.Add(""))
            Me.Invoke(Sub() newItem.SubItems.Add(""))
            Me.Invoke(Sub() newItem.SubItems.Add("Processing"))
        End If
        Me.Invoke(Sub() ListView1.EndUpdate())
    End Sub
    Private Sub updateFinishedListViewItem(itemIndex As Integer, outputFilePath As String)

        Dim currentItem As ListViewItem
        Me.Invoke(Sub() currentItem = ListView1.Items.Item(itemIndex))

        Dim currentFileInfo As New FileInfo(outputFilePath)
        Dim newItem As New ListViewItem
        newItem.Text = currentItem.Text
        newItem.Tag = currentItem.Tag

        newItem.SubItems.Add(currentItem.SubItems.Item(1).Text)
        newItem.SubItems.Item(1).Tag = currentItem.SubItems.Item(1).Tag

        newItem.SubItems.Add(convertBytesToAppropriateScale(currentFileInfo.Length))
        newItem.SubItems.Item(2).Tag = currentFileInfo.Length

        'Dim spaceSavedPercentage = Math.Round((Val(currentItem.SubItems.Item(1).Tag) / currentFileInfo.Length) * 100 - 100, 2)
        Dim spaceSavedPercentage = Math.Round(((Val(currentItem.SubItems.Item(1).Tag) - currentFileInfo.Length) / Val(currentItem.SubItems.Item(1).Tag)) * 100, 2)

        newItem.SubItems.Add(spaceSavedPercentage.ToString + "%")
        newItem.SubItems.Add("Done")

        Me.Invoke(Sub() ListView1.Items.Item(itemIndex) = newItem)

    End Sub
    Private Function getNonExistingPath(inputPath As String, extention As String)
        Dim count = 0
        While My.Computer.FileSystem.FileExists(inputPath + extention)
            If Not My.Computer.FileSystem.FileExists(String.Format("{0}{1}{2}", inputPath, count, extention)) Then
                inputPath = String.Format("{0}{1}", inputPath, count)
            End If
        End While
        Return inputPath + extention
    End Function
#End Region

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize
    End Sub

    Private Sub ListView1_DragEnter(sender As Object, e As DragEventArgs) Handles ListView1.DragEnter
        If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub ListView1_DragDrop(sender As Object, e As DragEventArgs) Handles ListView1.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)
        For Each path In files
            addFileToList(path)
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        optimiseImages()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadSettings()
        'Double buffer listview to prevent flickering
        Dim controlProperty As System.Reflection.PropertyInfo = GetType(System.Windows.Forms.Control).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.NonPublic Or System.Reflection.BindingFlags.Instance)
        controlProperty.SetValue(ListView1, True, Nothing)

        'Cannot be set to less threads than is available on the CPU or it wont work. :(
        Dim setMaxThreadsSuccess = ThreadPool.SetMaxThreads(Environment.ProcessorCount, Environment.ProcessorCount)
        If Not setMaxThreadsSuccess Then
            MsgBox("Error setting max threads")
        End If
    End Sub

    Private Sub ListView1_KeyDown(sender As Object, e As KeyEventArgs) Handles ListView1.KeyDown
        If e.KeyCode = Keys.Delete Then
            ListView1.BeginUpdate()
            For i As Integer = ListView1.SelectedIndices.Count - 1 To 0 Step -1
                ListView1.Items.RemoveAt(ListView1.SelectedIndices(i))
            Next
            ListView1.EndUpdate()
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click
        Dim fbd As New FolderBrowserDialog
        If fbd.ShowDialog() = Windows.Forms.DialogResult.OK Then
            If Not fbd.SelectedPath.Substring(fbd.SelectedPath.Length - 1, 1) = "\" Then
                saveLocation = fbd.SelectedPath + "\"
            Else
                saveLocation = fbd.SelectedPath
            End If
        End If
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem3.Click
        Dim ofd As New OpenFileDialog
        ofd.Filter = "All images|*.jpg;*.jpeg"
        ofd.Multiselect = True
        If ofd.ShowDialog = Windows.Forms.DialogResult.OK Then
            For Each filePath In ofd.FileNames
                addFileToList(filePath)
            Next
        End If
    End Sub

    Private Sub OptionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OptionsToolStripMenuItem.Click
        Form2.Show()
    End Sub

    Private Sub ListView1_ClientSizeChanged(sender As Object, e As EventArgs) Handles ListView1.ClientSizeChanged

        Dim otherColumnsWith As Integer = 0
        For i = 1 To ListView1.Columns.Count - 1
            otherColumnsWith += ListView1.Columns.Item(i).Width
        Next
        If visibleScrollbars.IsVScrollVisible(ListView1) Then
            ListView1.Columns.Item(0).Width = ListView1.Width - otherColumnsWith - 22
        Else
            ListView1.Columns.Item(0).Width = ListView1.Width - otherColumnsWith - 4
        End If
        If visibleScrollbars.IsHScrollVisible(ListView1) Then
            ListView1.Update()
        End If
    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        saveSettings()
    End Sub

    Private Sub ListView1_DoubleClick(sender As Object, e As EventArgs) Handles ListView1.DoubleClick
        Process.Start(ListView1.SelectedItems(0).Tag)
    End Sub
End Class
