Imports System.IO
Public Class SettingsFile
    Public Shared Sub saveSettings(settingsDict As Dictionary(Of String, String), filePath As String)
        Dim outFile As New StreamWriter(filePath)
        For Each item In settingsDict.Keys
            outFile.WriteLine(String.Format("{0}={1}", item, settingsDict(item)))
        Next
        outFile.Close()
    End Sub

    Public Shared Function loadSettings(filePath As String)
        Dim inFile As New StreamReader(filePath)
        Dim settingsDict As New Dictionary(Of String, String)

        Dim rtb As New RichTextBox
        rtb.Text = inFile.ReadToEnd

        For Each line As String In rtb.Lines
            If line.Contains("=") Then
                Dim setting = line.Split("=")
                settingsDict.Add(setting(0), setting(1))
            End If
        Next
        inFile.Close()
        Return settingsDict
    End Function
End Class
