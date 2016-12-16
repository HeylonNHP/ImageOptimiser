﻿Imports System.IO
Imports System.Threading

Public Class Form1
#Region "Basic functions"
    Private Function convertBytesToAppropriateScale(bytes As Integer, Optional precision As Integer = 2)
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
#End Region

#Region "Optimise"
    Public Const JPEGtranLocation As String = "C:\Users\Heylon2\PycharmProjects\JPEGtran_optimiser\jpegtran.exe"
    Private Sub optimiseImages()

        For Each item As ListViewItem In ListView1.Items
            Dim currentFile As New FileInfo(item.Tag)

            'Get current path without extention
            Dim outputFilename As String = currentFile.FullName.Substring(0, currentFile.FullName.Length - currentFile.Extension.Length)

            'Detirmine file type and apply appropriate optimisation
            If currentFile.Extension.ToLower = ".jpg" Or currentFile.Extension.ToLower = ".jpeg" Then
                Dim jpegOptimWCB As WaitCallback = New WaitCallback(AddressOf optimiseJPEGthread)
                ThreadPool.QueueUserWorkItem(jpegOptimWCB, {currentFile.FullName, outputFilename})
            End If
        Next
    End Sub

    Private Sub optimiseJPEGthread(parameters As Array)
        optimiseJPEG(parameters(0), parameters(1))
    End Sub
    Private Sub optimiseJPEG(inputFilePath As String, outputFilePath As String)
        Dim inputFileInfo As New FileInfo(inputFilePath)
        outputFilePath = getNonExistingPath(outputFilePath, inputFileInfo.Extension)

        updatedProcessingListViewItem(inputFileInfo.FullName)

        Dim JPEGtranProcess As New Process

        With JPEGtranProcess.StartInfo
            .FileName = JPEGtranLocation
            .Arguments = String.Format("-optimize -progressive -outfile {0} {1}", _
                                      """" + outputFilePath + """", """" + inputFilePath + """")
            .UseShellExecute = False
            .CreateNoWindow = True
        End With

        JPEGtranProcess.Start()

        JPEGtranProcess.PriorityClass = ProcessPriorityClass.Idle

        While Not JPEGtranProcess.HasExited
            Thread.Sleep(100)
        End While

        Try
            My.Computer.FileSystem.DeleteFile(inputFilePath)
            My.Computer.FileSystem.RenameFile(outputFilePath, inputFileInfo.Name)
        Catch ex As Exception

        End Try

        updateFinishedListViewItem(inputFilePath)
        Debug.Print(inputFilePath + " Ended")
    End Sub

    Private Sub updatedProcessingListViewItem(filePath As String)
        For i As Integer = 0 To ListView1.Items.Count - 1
            Dim currentItem As ListViewItem
            Me.Invoke(Sub() currentItem = ListView1.Items.Item(i))

            If currentItem.Tag = filePath Then
                Dim newItem As ListViewItem = currentItem
                Me.Invoke(Sub() newItem.SubItems.Add(""))
                Me.Invoke(Sub() newItem.SubItems.Add(""))
                Me.Invoke(Sub() newItem.SubItems.Add("Processing"))
            End If
        Next
    End Sub
    Private Sub updateFinishedListViewItem(filePath As String)
        For i As Integer = 0 To ListView1.Items.Count - 1
            Dim currentItem As ListViewItem
            Me.Invoke(Sub() currentItem = ListView1.Items.Item(i))

            If currentItem.Tag = filePath Then
                Dim currentFileInfo As New FileInfo(filePath)
                Dim newItem As New ListViewItem
                newItem.Text = currentItem.Text
                newItem.Tag = currentItem.Tag

                newItem.SubItems.Add(currentItem.SubItems.Item(1).Text)
                newItem.SubItems.Item(1).Tag = currentItem.SubItems.Item(1).Tag

                newItem.SubItems.Add(convertBytesToAppropriateScale(currentFileInfo.Length))
                newItem.SubItems.Item(2).Tag = currentFileInfo.Length

                Dim spaceSavedPercentage = Math.Round((Val(currentItem.SubItems.Item(1).Tag) / currentFileInfo.Length) * 100 - 100, 2)

                newItem.SubItems.Add(spaceSavedPercentage.ToString + "%")
                newItem.SubItems.Add("Done")

                Me.Invoke(Sub() ListView1.Items.Item(i) = newItem)
            End If
        Next
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

        Dim otherColumnsWith As Integer = 0
        For i = 1 To ListView1.Columns.Count - 1
            otherColumnsWith += ListView1.Columns.Item(i).Width
        Next
        ListView1.Columns.Item(0).Width = ListView1.Width - otherColumnsWith - 4
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
            Dim currentFile As New FileInfo(path)
            Dim newItem As New ListViewItem
            newItem.Text = currentFile.Name
            newItem.Tag = currentFile.FullName
            newItem.SubItems.Add(convertBytesToAppropriateScale(currentFile.Length))
            newItem.SubItems.Item(1).Tag = currentFile.Length
            ListView1.Items.Add(newItem)
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        optimiseImages()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Cannot be set to less threads than is available on the CPU or it wont work. :(
        ThreadPool.SetMaxThreads(4, 4)
    End Sub

    Private Sub ListView1_KeyDown(sender As Object, e As KeyEventArgs) Handles ListView1.KeyDown
        If e.KeyCode = Keys.Delete Then
            For i As Integer = ListView1.SelectedIndices.Count - 1 To 0 Step -1
                ListView1.Items.RemoveAt(ListView1.SelectedIndices(i))
            Next
        End If
    End Sub
End Class
