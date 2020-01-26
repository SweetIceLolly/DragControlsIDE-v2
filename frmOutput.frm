VERSION 5.00
Begin VB.Form frmOutput 
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "输出"
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
'描述:      输出窗口
'作者:      冰棍
'文件:      frmOutput.frm
'====================================================

Option Explicit

'输出行所对应的文件信息
Private Type OutputLineInfo
    InfoType        As Boolean              '操作类型（False: 在代码窗口显示; True: 在文件浏览器中显示FileName指定的路径，此时FileLine无意义）
    LineIndex       As Long                 '行序号
    FileName        As String               '文件名
    FileLine        As Long                 '文件行号
    FileColumn      As Long                 '文件列号
End Type

Private LineInfo()  As OutputLineInfo

'描述:      清空所有东西，为下一次调试做准备
Public Sub ClearEverything()
    Me.edOutput.Text = ""
    ReDim LineInfo(0)
End Sub

'描述:      为当前文本框中的最后一行添加输出行信息到LineInfo中
'参数:      FileName: 该行对应的文件名
'.          FileLine: 该行对应的文件行号
'           FileColumn: 该行对应的文件列号
Public Sub AddLineInfo(InfoType As Boolean, FileName As String, FileLine As Long, Optional FileColumn As Long = -1)
    Dim tmpInfo     As OutputLineInfo
    
    tmpInfo.InfoType = InfoType
    tmpInfo.LineIndex = SendMessageA(Me.edOutput.hwnd, EM_GETLINECOUNT, ByVal 0, ByVal 0) - 2       '把对应的输出行的序号记录下来
    tmpInfo.FileName = FileName
    tmpInfo.FileLine = FileLine
    tmpInfo.FileColumn = FileColumn
    
    LineInfo(UBound(LineInfo)) = tmpInfo
    ReDim Preserve LineInfo(UBound(LineInfo) + 1)
End Sub

'描述:      往窗体的文本框上输出带时间的消息
'参数:      strOutput: 需要输出的消息
Public Sub OutputLog(strOutput As String)
    Me.edOutput.Text = Me.edOutput.Text & Date & " " & Time & " " & strOutput & vbCrLf
    Me.edOutput.SelStart = Len(Me.edOutput.Text)                                            '滚动到文本末尾
End Sub

Private Sub edOutput_DblClick()
    On Error Resume Next
    Dim LineIndex   As Long                 '当前光标在输出里对应的行号
    Dim FileIndex   As Long                 '当前行所对应的代码文件在CurrentProject.Files中的序号
    Dim i           As Long
    
    LineIndex = SendMessageA(Me.edOutput.hwnd, EM_LINEFROMCHAR, ByVal -1, ByVal 0)          '获取光标所在行
    For i = 0 To UBound(LineInfo) - 1
        If LineInfo(i).LineIndex = LineIndex Then
            If LineInfo(i).InfoType = False Then                                                '执行“在代码窗口中显示”操作
                Dim NewCodeWindow       As frmCodeWindow
                
                Set NewCodeWindow = frmMain.ShowCodeWindow(, LineInfo(i).FileName)                  '切换到该文件对应的代码窗口
                If NewCodeWindow Is Nothing Then
                    NoSkinMsgBox Lang_Main_Debug_OpenSourceFailure & LineInfo(i).FileName, vbExclamation, Lang_Msgbox_Error
                Else
                    NewCodeWindow.SyntaxEdit.CurrPos.Row = LineInfo(i).FileLine
                    If LineInfo(i).FileColumn <> -1 And NewCodeWindow.SyntaxEdit.TabWithSpace = True Then
                        NewCodeWindow.SyntaxEdit.CurrPos.Col = LineInfo(i).FileColumn
                    End If
                    NewCodeWindow.SyntaxEdit.SetFocus
                End If
            Else                                                                                '执行“在文件浏览器中显示”操作
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
