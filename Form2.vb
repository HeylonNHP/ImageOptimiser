Public Class Form2

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Process priority
        If Form1.processPriority = ProcessPriorityClass.Idle Then
            ComboBox1.SelectedIndex = 5
        ElseIf Form1.processPriority = ProcessPriorityClass.BelowNormal Then
            ComboBox1.SelectedIndex = 4
        ElseIf Form1.processPriority = ProcessPriorityClass.Normal Then
            ComboBox1.SelectedIndex = 3
        ElseIf Form1.processPriority = ProcessPriorityClass.AboveNormal Then
            ComboBox1.SelectedIndex = 2
        ElseIf Form1.processPriority = ProcessPriorityClass.High Then
            ComboBox1.SelectedIndex = 1
        ElseIf Form1.processPriority = ProcessPriorityClass.RealTime Then
            ComboBox1.SelectedIndex = 0
        End If

        'copy exif
        If Form1.copyExif = 0 Then
            RadioButton1.Checked = True
        ElseIf Form1.copyExif = 1 Then
            RadioButton2.Checked = True
        ElseIf Form1.copyExif = 2 Then
            RadioButton3.Checked = True
        End If
        'huffman table
        CheckBox1.Checked = Form1.optimiseHuffmanTable
        'progressive
        CheckBox2.Checked = Form1.convertToProgressive
        'arithmetic
        CheckBox3.Checked = Form1.arithmeticCoding
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Process priority
        If ComboBox1.SelectedIndex = 0 Then
            Form1.processPriority = ProcessPriorityClass.RealTime
        ElseIf ComboBox1.SelectedIndex = 1 Then
            Form1.processPriority = ProcessPriorityClass.High
        ElseIf ComboBox1.SelectedIndex = 2 Then
            Form1.processPriority = ProcessPriorityClass.AboveNormal
        ElseIf ComboBox1.SelectedIndex = 3 Then
            Form1.processPriority = ProcessPriorityClass.Normal
        ElseIf ComboBox1.SelectedIndex = 4 Then
            Form1.processPriority = ProcessPriorityClass.BelowNormal
        ElseIf ComboBox1.SelectedIndex = 5 Then
            Form1.processPriority = ProcessPriorityClass.Idle
        End If
        'copy exif
        If RadioButton1.Checked Then
            Form1.copyExif = 0
        ElseIf RadioButton2.Checked Then
            Form1.copyExif = 1
        ElseIf RadioButton3.Checked Then
            Form1.copyExif = 2
        End If

        'huffman table
        Form1.optimiseHuffmanTable = CheckBox1.Checked
        'progressive
        Form1.convertToProgressive = CheckBox2.Checked
        'arithmetic
        Form1.arithmeticCoding = CheckBox3.Checked
        Me.Close()
    End Sub
End Class