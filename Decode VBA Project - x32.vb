'From: http://stackoverflow.com/questions/1026483/is-there-a-way-to-crack-the-password-on-an-excel-vba-project/31005696#31005696

''Module 1

Option Explicit

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" _
         (Destination As Long, Source As Long, ByVal Length As Long)

Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Long, _
         ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, _
         ByVal lpProcName As String) As Long
Private Declare Function DialogBoxParam Lib "user32" Alias "DialogBoxParamA" (ByVal hInstance As Long, _
         ByVal pTemplateName As Long, ByVal hWndParent As Long, _
         ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer
Dim HookBytes(0 To 5) As Byte
Dim OriginBytes(0 To 5) As Byte
Dim pFunc As Long
Dim Flag As Boolean

Private Function GetPtr(ByVal Value As Long) As Long
     '获得函数的地址
     GetPtr = Value
End Function

Public Sub RecoverBytes()
     '若已经hook,则恢复原API开头的6字节,也就是恢复原来函数的功能
     If Flag Then MoveMemory ByVal pFunc, ByVal VarPtr(OriginBytes(0)), 6
End Sub

Public Function Hook() As Boolean
     Dim TmpBytes(0 To 5) As Byte
     Dim p As Long
     Dim OriginProtect As Long
     Hook = False
     'VBE6.dll调用DialogBoxParamA显示VB6INTL.dll资源中的第4070号对话框(就是输入密码的窗口)
     '若DialogBoxParamA返回值非0,则VBE会认为密码正确,所以我们要hook DialogBoxParamA函数
     pFunc = GetProcAddress(GetModuleHandleA("user32.dll"), "DialogBoxParamA")
     '标准api hook过程之一: 修改内存属性,使其可写
     If VirtualProtect(ByVal pFunc, 6, &H40, OriginProtect) <> 0 Then
         '标准api hook过程之二: 判断是否已经hook,看看API的第一个字节是否为&H68,
         '若是则说明已经Hook
         MoveMemory ByVal VarPtr(TmpBytes(0)), ByVal pFunc, 6
         If TmpBytes(0) <> &H68 Then
             '标准api hook过程之三: 保存原函数开头字节,这里是6个字节,以备后面恢复
             MoveMemory ByVal VarPtr(OriginBytes(0)), ByVal pFunc, 6
             '用AddressOf获取MyDialogBoxParam的地址
             '因为语法不允许写成p = AddressOf MyDialogBoxParam,这里我们写一个函数
             'GetPtr,作用仅仅是返回AddressOf MyDialogBoxParam的值,从而实现将
             'MyDialogBoxParam的地址付给p的目的
             p = GetPtr(AddressOf MyDialogBoxParam)
             '标准api hook过程之四: 组装API入口的新代码
             'HookBytes 组成如下汇编
             'push MyDialogBoxParam的地址
             'ret
             '作用是跳转到MyDialogBoxParam函数
             HookBytes(0) = &H68
             MoveMemory ByVal VarPtr(HookBytes(1)), ByVal VarPtr(p), 4
             HookBytes(5) = &HC3
             '标准api hook过程之五: 用HookBytes的内容改写API前6个字节
             MoveMemory ByVal pFunc, ByVal VarPtr(HookBytes(0)), 6
             '设置hook成功标志
             Flag = True
             Hook = True
         End If
     End If
End Function

Private Function MyDialogBoxParam(ByVal hInstance As Long, _
         ByVal pTemplateName As Long, ByVal hWndParent As Long, _
         ByVal lpDialogFunc As Long, ByVal dwInitParam As Long) As Integer
     If pTemplateName = 4070 Then
         '有程序调用DialogBoxParamA装入4070号对话框,这里我们直接返回1,让
         'VBE以为密码正确了
         MyDialogBoxParam = 1
     Else
         '有程序调用DialogBoxParamA,但装入的不是4070号对话框,这里我们调用
         'RecoverBytes函数恢复原来函数的功能,在进行原来的函数
         RecoverBytes
         MyDialogBoxParam = DialogBoxParam(hInstance, pTemplateName, _
                            hWndParent, lpDialogFunc, dwInitParam)
         '原来的函数执行完毕,再次hook
         Hook
     End If
End Function

''Module 2 or onwrite in a worksheet

Sub Decode()

If Hook Then

MsgBox "Decode successfully"

End If

End Sub

Sub Recovery()

RecoverBytes

MsgBox "Recovery successfully"

End Sub