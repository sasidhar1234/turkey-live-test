Public Class Form2
    Dim myexcel As Microsoft.Office.Interop.Excel.Application
    Dim myworkbook, myworkbook_AI As Microsoft.Office.Interop.Excel.Workbook
    Dim myoutputarray(,) As String
    Dim myworksheet, myworksheet_AI As Microsoft.Office.Interop.Excel.Worksheet
    Dim niterations As Long

    Public Sub initate_excel()
        Try
            myexcel = New Microsoft.Office.Interop.Excel.Application
            myworkbook = myexcel.Workbooks.Open("C:\Codes\Live turkey Movement\AIoutput\Multi run.xlsx")
        Catch ex As Exception
            myexcel.Quit()
            MsgBox(ex)
        End Try
    End Sub
    Public Sub readinputandinitiatearrays()
        myworksheet = myworkbook.Worksheets("Input")
        niterations = myworksheet.Range("B3").Value
        ReDim myoutputarray(niterations - 1, 6)
        Dim i As Integer
        For i = 0 To (niterations - 1)
            myoutputarray(i, 0) = i
            myoutputarray(i, 1) = myworksheet.Range("C" & (4 + i)).Value
        Next
    End Sub
    Public Sub iterate(iter_no As Integer)
        form1.Show()
        ' form1.Givendiseasedayonmovment = myoutputarray(iter_no, 1) '
        form1.survsimparametertochange = myoutputarray(iter_no, 1)
        'the above coud be form1.mysurvsimparameter to change any survsim parameter
        form1.dummybutton2_click()
        initate_excel_AI()
        myworksheet_AI = myworkbook_AI.Worksheets("survpar")
        myoutputarray(iter_no, 2) = myworksheet_AI.Range("BA3").Value
        myoutputarray(iter_no, 3) = myworksheet_AI.Range("BB3").Value
        myoutputarray(iter_no, 4) = myworksheet_AI.Range("BC3").Value
        myoutputarray(iter_no, 5) = myworksheet_AI.Range("BD3").Value
        'last is the detection percent with infectious birds
        myoutputarray(iter_no, 6) = myworksheet_AI.Range("AZ3").Value
        close_excel_AI()
        form1.Close()
    End Sub
    Public Sub writeoutput()
        myworksheet = myworkbook.Worksheets("Output")
        Dim i, k As Integer
        For i = 0 To niterations - 1
            For k = 0 To 6
                myworksheet.Cells(3 + i, 3 + k) = myoutputarray(i, k)
            Next
        Next
    End Sub

    Public Sub initate_excel_AI()
        Try
            myworkbook_AI = myexcel.Workbooks.Open("C:\Codes\Live turkey Movement\AIoutput\AI Output.xlsx")
        Catch ex As Exception
            myexcel.Quit()
            MsgBox(ex)
        End Try
    End Sub
    Public Sub close_excel()
        myworkbook.Close(True)
        myexcel.Quit()
    End Sub
    Public Sub close_excel_AI()
        myworkbook_AI.Close(True)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        form1.Close()
        form1.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        form1.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        initate_excel()
        readinputandinitiatearrays()
        Dim k As Integer
        For k = 0 To niterations - 1
            iterate(k)
            TextBox1.Text = k
            Me.Show()
        Next
        writeoutput()
        close_excel()
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class