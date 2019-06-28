Attribute VB_Name = "modConfig"
'====================================================
'描述:      提供读写程序配置文件，包括程序设置、语言、用户习惯等函数
'作者:      冰棍
'文件:      modConfig.bas
'====================================================

Option Explicit

'描述:      读取对应语言的字符串资源。该函数会通过
'.          提供的第一个资源ID来计算出其他字符串所对应的ID
'参数:      ResID: 对应语言所对应的第一个资源ID。如本程序中文语言所对应的第一个资源ID是1001
'返回值:    如果读取成功，返回True；否则返回False
Public Function LoadLanguage(ResID As Long) As Boolean
    On Error Resume Next
    LoadLanguage = True
    
    '读取菜单字符串
    Dim id          As Long
    
    For id = 0 To 69
        frmMain.DarkMenu.MenuText(id) = LoadResString(ResID + id)
        If Err.Number <> 0 Then
            LoadLanguage = False
            Exit Function
        End If
    Next id
End Function
