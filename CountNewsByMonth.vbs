Sub CountNewsByMonth()

    Dim currentMonth As Integer
    Dim currentCount As Integer
    Dim lastRow As Long
    Dim currentDate As Date
    
    currentMonth = 0
    currentCount = 0
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 1 To lastRow
        currentDate = Cells(i, "A").Value
        If Month(currentDate) <> currentMonth Then
            If currentMonth <> 0 Then
                Cells(i - 1, "B").Value = currentCount
            End If
            currentMonth = Month(currentDate)
            currentCount = 1
        Else
            currentCount = currentCount + 1
        End If
    Next i
    
    Cells(lastRow, "B").Value = currentCount

End Sub
