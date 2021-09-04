'Create a index for all worksheets, wrote in 2012

Sub G_Index()
    On Error Resume Next
    Application.ScreenUpdating = False
    
    'Create index sheet name
    ShN = InputBox("Please input the name of the Index", "INPUT", "Index")
    If ShN = "" Then
        Exit Sub
    Else
        For Each Sheet In Sheets
            If Sheet.Name = ShN Then
                response1 = MsgBox("The name you inputed is Duplicated or Unkown Mistakes !", _
                vbOKOnly + vbInformation, "Warning")
            Exit Sub
            End If
         Next
    End If
    
    'Add reference to other sheets in the index sheet
    Sheets.Add
    ActiveSheet.Move Before:=Sheets(1)
    ActiveSheet.Name = ShN
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "Index of Worksheets"
    Columns("B:B").EntireColumn.AutoFit
    Selection.Font.Bold = True
    Range("B3").Select
        For Each Sheet In Sheets
            ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
                "'" & Sheet.Name & "'" & "!a1", TextToDisplay:=Sheet.Name, ScreenTip:=Sheet.Name
            ActiveCell.Offset(1, 0).Range("A1").Select
        Next
    
    Range("B3").Delete
    Range("B2").Select

    Application.ScreenUpdating = True
    
End Sub