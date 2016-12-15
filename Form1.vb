Public Class Form1

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles MyBase.Resize

        Dim otherColumnsWith As Integer = 0
        For i = 1 To ListView1.Columns.Count - 1
            otherColumnsWith += ListView1.Columns.Item(i).Width
        Next
        ListView1.Columns.Item(0).Width = ListView1.Width - otherColumnsWith - 4
    End Sub
End Class
