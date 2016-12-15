Imports System.IO
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
        Me.Invoke(Sub() ListView1.Items.Add("Test threading"))
    End Sub
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
            ListView1.Items.Add(newItem)
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        optimiseImages()
    End Sub
End Class
