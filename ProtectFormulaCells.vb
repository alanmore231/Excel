'Requsted by Erica and wrote in 2015. Can easily change password or to protect other functions.

'Protect all formula in active workbook
Sub protectformula()

Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Sheets
        ws.Unprotect Password = "Passwd"
        ws.Cells.Locked = False
        maxrow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        maxcol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
        For i = 1 To maxrow
            For j = 1 To maxcol
                If ws.Cells(i, j).HasFormula Then
                    ws.Cells(i, j).Locked = True
                End If
            Next
        Next
        ws.Protect Password = "Passwd"
    Next
End Sub