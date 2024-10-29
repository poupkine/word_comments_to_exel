Sub CopyCommentsToExcel()
'Create in Word vba
Dim xlApp As Object
Dim xlWB As Object
Dim i As Integer
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If Err Then
        Set xlApp = CreateObject("Excel.Application")
    End If
    On Error GoTo 0
    xlApp.Visible = True
    Set xlWB = xlApp.Workbooks.Add        ' create a new workbook
    With xlWB.Worksheets(1)
        For i = 1 To ActiveDocument.Comments.Count
            .Cells(i, 1).Formula = ActiveDocument.Comments(i).Author
            .Cells(i, 2).Formula = ActiveDocument.Comments(i).Range
            .Cells(i, 3).Formula = ActiveDocument.Comments(i).Scope.Text
            .Cells(i, 4).Formula = Format(ActiveDocument.Comments(i).Date, "dd/MM/yyyy")
            
        Next i
    End With
    Set xlWB = Nothing
    Set xlApp = Nothing
End Sub
