Attribute VB_Name = "Módulo1"
Sub sorter()
    Dim i As Integer, j As Integer, temp As Integer, rng As Range
    Set rng = Range("A1").CurrentRegion
    For i = 1 To rng.Count
        For j = i + 1 To rng.Count
            If rng.Cells(j) < rng.Cells(i) Then
                'swap numbers
                temp = rng.Cells(i)
                rng.Cells(i) = rng.Cells(j)
                rng.Cells(j) = temp
            End If
        Next j
    Next i
End Sub
