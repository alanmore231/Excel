# From: https://archive.codeplex.com/?p=removestyles 
# Replace by official addin: https://support.microsoft.com/en-us/office/clean-excess-cell-formatting-on-a-worksheet-e744c248-6925-4e77-9d49-4874f7474738

## Workbook (enable ribbon)

Dim WithEvents app As Application

Private Sub app_WorkbookActivate(ByVal Wb As Workbook)
    Module1.MyRibbon.Invalidate
End Sub

Private Sub Workbook_Open()
    Set app = Application
End Sub

## Module 1

Public MyRibbon As IRibbonUI
'Callback for customUI.onLoad
Sub CallbackOnLoad(ribbon As IRibbonUI)
    Set MyRibbon = ribbon
End Sub

'Callback for customButton getLabel
Sub GetButtonLabel(control As IRibbonControl, ByRef returnedVal)
    If ActiveWorkbook Is Nothing Then
        returnedVal = "Remove Styles"
    Else
        returnedVal = "Remove Styles" & vbCr & Format(ActiveWorkbook.Styles.Count, "#" & Application.International(xlThousandsSeparator) & "##0")
    End If
End Sub


Sub RemoveTheStyles(control As IRibbonControl)
    Dim s As Style, i As Long, c As Long
    On Error Resume Next
    If ActiveWorkbook.MultiUserEditing Then
        If MsgBox("You cannot remove Styles in a Shared workbook." & vbCr & vbCr & _
                  "Do you want to unshare the workbook?", vbYesNo + vbInformation) = vbYes Then
            ActiveWorkbook.ExclusiveAccess
            If Err.Description = "Application-defined or object-defined error" Then
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If
    c = ActiveWorkbook.Styles.Count
    Application.ScreenUpdating = False
    For i = c To 1 Step -1
        If i Mod 600 = 0 Then DoEvents
        Set s = ActiveWorkbook.Styles(i)
        Application.StatusBar = "Deleting " & c - i + 1 & " of " & c & " " & s.Name
        If Not s.BuiltIn Then
            s.Delete
            If Err.Description = "You cannot use this command on a protected sheet. To use this command, you must first unprotect the sheet (Review tab, Changes group, Unprotect Sheet button). You may be prompted for a password." Then
                MsgBox Err.Description & vbCr & "You have to unprotect all of the sheets in the workbook to remove styles.", vbExclamation, "Remove Styles AddIn"
                Exit For
            End If
        End If
    Next
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub
