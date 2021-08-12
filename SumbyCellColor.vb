'Sum by cell color, a function written in 2015

Function ColorSum(Ref_color As Range, Sum_range As Range) 

    Application.Volatile
    Dim iCol As Integer
    Dim rCell As Range

    iCol = Ref_color.Interior.ColorIndex
    
    For Each rCell In Sum_range
        If iCol = rCell.Interior.ColorIndex Then
            ColorSum = ColorSum + rCell.Value
        End If
    Next rCell
    
End Function