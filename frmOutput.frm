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
'����:      �������
'����:      ����
'�ļ�:      frmOutput.frm
'====================================================

Option Explicit

'����:      ��������ı����������ʱ�����Ϣ
'����:      strOutput: ��Ҫ�������Ϣ
Public Sub OutputLog(strOutput As String)
    Me.edOutput.Text = Me.edOutput.Text & Date & " " & Time & " " & strOutput & vbCrLf
    Me.edOutput.SelStart = Len(Me.edOutput.Text)                                            '�������ı�ĩβ
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.edOutput.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
