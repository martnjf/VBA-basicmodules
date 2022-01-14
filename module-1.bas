Attribute VB_Name = "Módulo1"
Sub ocultarfilas()
    'Dim StartTime As Double
    'Dim SecondsElapsed As Double
    'StartTime = Timer
    '*****************************
    StartRow = 7
    EndRow = 102
    ColNum = 8
    ColNumB = 9
    For i = StartRow To EndRow
    If (Cells(i, ColNum).Value = 2) And (Cells(i, ColNumB).Value = 2) Then
        Cells(i, ColNum).EntireRow.Hidden = True
    ElseIf (Cells(i, ColNum).Value = 2) And (Cells(i, ColNumB).Value = 1) Then
        Cells(i, ColNum).EntireRow.Hidden = True
    ElseIf (Cells(i, ColNum).Value = 1) And (Cells(i, ColNumB).Value = 2) Then
        Cells(i, ColNum).EntireRow.Hidden = True
    Else
        Cells(i, ColNum).EntireRow.Hidden = False
    End If
    Next i

    'SecondsElapsed = Round(Timer - StartTime, 2)
    'MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation

End Sub
