VERSION 5.00
Begin VB.Form frmStartupLogo 
   Appearance      =   0  'Flat
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "�Ͽؼ���"
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   Picture         =   "frmStartupLogo.frx":0000
   ScaleHeight     =   1590
   ScaleWidth      =   5445
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmStartupLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================
'����:      ��������
'����:      Error 404
'�ļ�:      frmStartupLogo.frm
'====================================================

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    '���ڴ�С��Ӧ��ͬDPI
    Me.Width = 533 * Screen.TwipsPerPixelX
    Me.Height = 160 * Screen.TwipsPerPixelY
End Sub
