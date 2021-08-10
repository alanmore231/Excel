# Requested by Emily and wrote in 2017

Sub Find()
 Application.ScreenUpdating = False

 Dim MyDir As String
 MyDir = ThisWorkbook.Path & "\"
 ChDrive Left(MyDir, 1) 'find all the excel files
 ChDir MyDir
 Match = Dir$("")

 i = 1
 Do
 If Not LCase(Match) = LCase(ThisWorkbook.Name) Then
    Workbooks.Open Match, 0 'open
    ActiveWorkbook.Sheets("sheetneedtocopy").Select
    ActiveSheet.Name = Left(ActiveWorkbook.Name, Application.Find(".", ActiveWorkbook.Name) - 1)
    ActiveSheet.Copy Before:=ThisWorkbook.Sheets(1) 'copy sheet
    i = i + 1
    Windows(Match).Activate
    ActiveWindow.Close Savechanges:=False
 End If

 Match = Dir$
 Loop Until Len(Match) = 0

 Application.ScreenUpdating = True
 
 End Sub