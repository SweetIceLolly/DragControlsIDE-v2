Attribute VB_Name = "modConfig"
'====================================================
'����:      �ṩ��д���������ļ��������������á����ԡ��û�ϰ�ߵȺ���
'����:      ����
'�ļ�:      modConfig.bas
'====================================================

Option Explicit

'����:      ��ȡ��Ӧ���Ե��ַ�����Դ���ú�����ͨ��
'.          �ṩ�ĵ�һ����ԴID������������ַ�������Ӧ��ID
'����:      ResID: ��Ӧ��������Ӧ�ĵ�һ����ԴID���籾����������������Ӧ�ĵ�һ����ԴID��1001
'����ֵ:    �����ȡ�ɹ�������True�����򷵻�False
Public Function LoadLanguage(ResID As Long) As Boolean
    On Error Resume Next
    LoadLanguage = True
    
    '��ȡ�˵��ַ���
    Dim id          As Long
    
    For id = 0 To 69
        frmMain.DarkMenu.MenuText(id) = LoadResString(ResID + id)
        If Err.Number <> 0 Then
            LoadLanguage = False
            Exit Function
        End If
    Next id
End Function
