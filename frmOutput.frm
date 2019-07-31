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
      BorderStyle     =   0  'None
      ForeColor       =   &H00F1F1F1&
      Height          =   1695
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
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

'描述:      往窗体的文本框上输出带时间的消息
'参数:      strOutput: 需要输出的消息
Public Sub OutputLog(strOutput As String)
    Me.edOutput.Text = Me.edOutput.Text & Date & " " & Time & " " & strOutput & vbCrLf
    Me.edOutput.SelStart = Len(Me.edOutput.Text)                                            '滚动到文本末尾
End Sub

Private Sub Form_Load()
    Me.Caption = Lang_Output_Caption
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.edOutput.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
