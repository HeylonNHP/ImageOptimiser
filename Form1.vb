Imports System.IO
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
            newItem.SubItems.Add(convertBytesToAppropriateScale(currentFile.Length))
            ListView1.Items.Add(newItem)
        Next
    End Sub
End Class
