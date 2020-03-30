VERSION 5.00
Begin VB.Form frmOutput 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "���"
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox edOutput 
      Appearance      =   0  'Flat
      BackColor       =   &H00302D2D&
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F1F1F1&
      Height          =   1695
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "frmOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'����:      �������
'����:      ����
'�ļ�:      frmOutput.frm
'====================================================

Option Explicit

'���������Ӧ���ļ���Ϣ
Private Type OutputLineInfo
    InfoType        As Boolean              '�������ͣ�False: �ڴ��봰����ʾ; True: ���ļ����������ʾFileNameָ����·������ʱFileLine�����壩
    LineIndex       As Long                 '�����
    FileName        As String               '�ļ���
    FileLine        As Long                 '�ļ��к�
    FileColumn      As Long                 '�ļ��к�
End Type

Private LineInfo()  As OutputLineInfo

'����:      ������ж�����Ϊ��һ�ε�����׼��
Public Sub ClearEverything()
    Me.edOutput.Text = ""
    ReDim LineInfo(0)
End Sub

'����:      Ϊ��ǰ�ı����е����һ������������Ϣ��LineInfo��
'����:      InfoType: False: �ڴ��봰����ʾ; True: ���ļ����������ʾ
'.          FileName: ���ж�Ӧ���ļ���
'.          FileLine: ���ж�Ӧ���ļ��к�
'           FileColumn: ���ж�Ӧ���ļ��к�
Public Sub AddLineInfo(InfoType As Boolean, FileName As String, FileLine As Long, Optional FileColumn As Long = -1)
    Dim tmpInfo     As OutputLineInfo
    
    tmpInfo.InfoType = InfoType
    tmpInfo.LineIndex = SendMessageA(Me.edOutput.hwnd, EM_GETLINECOUNT, ByVal 0, ByVal 0) - 2       '�Ѷ�Ӧ������е���ż�¼����
    tmpInfo.FileName = FileName
    tmpInfo.FileLine = FileLine
    tmpInfo.FileColumn = FileColumn
    
    LineInfo(UBound(LineInfo)) = tmpInfo
    ReDim Preserve LineInfo(UBound(LineInfo) + 1)
End Sub

'����:      ��������ı����������ʱ�����Ϣ
'����:      strOutput: ��Ҫ�������Ϣ
Public Sub OutputLog(strOutput As String)
    Me.edOutput.Text = Me.edOutput.Text & Date & " " & Time & " " & strOutput & vbCrLf
    Me.edOutput.SelStart = Len(Me.edOutput.Text)                                            '�������ı�ĩβ
End Sub

Private Sub edOutput_DblClick()
    On Error Resume Next
    Dim LineIndex   As Long                 '��ǰ�����������Ӧ���к�
    Dim FileIndex   As Long                 '��ǰ������Ӧ�Ĵ����ļ���CurrentProject.Files�е����
    Dim i           As Long
    
    LineIndex = SendMessageA(Me.edOutput.hwnd, EM_LINEFROMCHAR, ByVal -1, ByVal 0)          '��ȡ���������
    For i = 0 To UBound(LineInfo) - 1
        If LineInfo(i).LineIndex = LineIndex Then
            If LineInfo(i).InfoType = False Then                                                'ִ�С��ڴ��봰������ʾ������
                Dim NewCodeWindow       As frmCodeWindow
                
                Set NewCodeWindow = frmMain.ShowCodeWindow(, LineInfo(i).FileName)                  '�л������ļ���Ӧ�Ĵ��봰��
                If NewCodeWindow Is Nothing Then
                    NoSkinMsgBox Lang_Main_Debug_OpenSourceFailure & LineInfo(i).FileName, vbExclamation, Lang_Msgbox_Error
                Else
                    NewCodeWindow.SyntaxEdit.CurrPos.Row = LineInfo(i).FileLine
                    If LineInfo(i).FileColumn <> -1 And NewCodeWindow.SyntaxEdit.TabWithSpace = True Then
                        NewCodeWindow.SyntaxEdit.CurrPos.Col = LineInfo(i).FileColumn
                    End If
                    NewCodeWindow.SyntaxEdit.SetFocus
                End If
            Else                                                                                'ִ�С����ļ����������ʾ������
                Shell "explorer.exe /select,""" & LineInfo(i).FileName & """", vbNormalFocus
            End If
            Exit Sub
        End If
    Next i
End Sub

Private Sub Form_Load()
    Me.Caption = Lang_Output_Caption
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.edOutput.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
