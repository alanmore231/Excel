'Requested by Patty and wrote in 2021, can easily extend to other executions that needs repeat in Word files in the same folder

Sub ReplaceWord()
Application.ScreenUpdating = False

Dim MyDir As String
Dim MatchFile As Variant
Dim OriginalText As String
Dim NewText As String

'Select folder
With Application.FileDialog(msoFileDialogFolderPicker)
    
    If .Show = -1 Then
        MyDir = .SelectedItems(1) & "\"
    Else
        Exit Sub
    End If
    
End With

'Set path
ChDrive Left(MyDir, 1)
ChDir MyDir
MatchFile = Dir$("")

'Get text need to replace
OriginalText = InputBox("Text needs to be replaced")
NewText = InputBox("Text after replace")


i = 1
Do
 If Not LCase(MatchFile) = LCase(ThisDocument.Name) Then
    
    Documents.Open MatchFile, 0

    'Below generate from record Macro for replacing text
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = OriginalText
        .Replacement.Text = NewText
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchByte = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    ActiveDocument.Save
    ActiveDocument.Close
        i = i + 1

 End If

 ChDrive Left(MyDir, 1)
 ChDir MyDir
 MatchFile = Dir$
 
Loop Until Len(MatchFile) = 0
 
Application.ScreenUpdating = True
 
End Sub