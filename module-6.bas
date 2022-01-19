Attribute VB_Name = "Módulo1"
Sub test()
    Dim wbk As Workbook
    Dim ws As Worksheet
    Dim i As Integer
    Dim count As Integer
    Set wbk = ThisWorkbook
    Set ws = wbk.Sheets(1)
    
    Dim cell As Range
    
    With ws
        For Each cel In ActiveSheet.Range("E1:E6")
            If cel.Font.Color <> RGB(255, 255, 0) Then
                Selection.Rows(cel).EntireRow.Delete
                'cel.Delete
                count = count + 1
            End If
        Next cel
    End With
    Cells(101, 5) = count
End Sub
