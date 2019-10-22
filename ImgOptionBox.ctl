VERSION 5.00
Object = "{ACD4732E-2B7C-40C1-A56B-078848D41977}#1.0#0"; "Image.ocx"
Begin VB.UserControl ImgOptionBox 
   BackColor       =   &H00303030&
   ClientHeight    =   3030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2685
   ScaleHeight     =   3030
   ScaleWidth      =   2685
   Begin VB.Label InputCover 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   2172
      Left            =   432
      TabIndex        =   1
      Top             =   336
      Width           =   1884
   End
   Begin VB.Shape focusBorder 
      BorderColor     =   &H00CEDB1A&
      BorderWidth     =   2
      Height          =   1860
      Left            =   528
      Top             =   576
      Visible         =   0   'False
      Width           =   1644
   End
   Begin VB.Label MyLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Console"
      ForeColor       =   &H00DEE2DE&
      Height          =   240
      Left            =   1008
      TabIndex        =   0
      Top             =   1944
      Width           =   720
   End
   Begin ImageX.aicAlphaImage MyIcon 
      Height          =   780
      Left            =   984
      Top             =   960
      Width           =   780
      _ExtentX        =   1376
      _ExtentY        =   1376
      Image           =   "ImgOptionBox.ctx":0000
   End
End
Attribute VB_Name = "ImgOptionBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================================
'描述:      ImageOptionBox控件
'作者:      Error 404
'文件:      ImgOptionBox.ctl
'====================================================

Event Click()

Dim imgData()   As Byte
Dim imgFileName As String
Dim IsFocus As Boolean

Public Sub ChangeAppearance(Mode As Boolean)
    If Mode Then
        focusBorder.Visible = True
        UserControl.BackColor = RGB(160, 160, 170)
        MyLabel.ForeColor = RGB(255, 255, 255)
        focusBorder.BorderColor = RGB(34, 151, 243)
    Else
        focusBorder.Visible = False
        UserControl.BackColor = RGB(100, 100, 105)
        MyLabel.ForeColor = RGB(180, 180, 180)
    End If
End Sub

Public Property Get Focused() As Boolean
    Focused = IsFocus
End Property

Public Property Let Focused(NewFocused As Boolean)
    If NewFocused Then
        Dim obj As Object
        For Each obj In UserControl.Parent.Controls
            If Not (obj Is Me) Then
                If TypeName(obj) = "ImgOptionBox" Then
                    obj.Focused = False
                End If
            End If
        Next
    End If
    IsFocus = NewFocused
    Call ChangeAppearance(IsFocus)
    
    RaiseEvent Click
End Property

Public Property Get Content() As String
    Content = MyLabel.Caption
End Property

Public Property Let Content(NewContent As String)
    MyLabel.Caption = NewContent
    MyLabel.Move UserControl.Width / 2 - MyLabel.Width / 2, UserControl.Height * 0.8 - MyLabel.Height / 2
    PropertyChanged "Content"
End Property

Public Property Get FileName() As String
    FileName = imgFileName
End Property

Public Property Let FileName(NewFileName As String)
    On Error Resume Next
    
    NewFileName = App.Path & "\icons\" & NewFileName
    
    imgFileName = NewFileName
    UserControl.MyIcon.LoadImage_FromFile NewFileName
    Open NewFileName For Binary As #1
        ReDim imgData(LOF(1))
        Get #1, , imgData
    Close #1
    
    PropertyChanged "Image"
End Property

Private Sub InputCover_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        If Not IsFocus Then
            Focused = True
        End If
        
        RaiseEvent Click
    End If
End Sub

Private Sub UserControl_Initialize()
    Call ChangeAppearance(False)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    imgData = PropBag.ReadProperty("Image", StrConv("", vbFromUnicode))
    
    MyLabel.Caption = PropBag.ReadProperty("Content", "")
    MyIcon.LoadImage_FromArray imgData
    
    Call UserControl_Resize
End Sub

Private Sub UserControl_Resize()
    focusBorder.Move Screen.TwipsPerPixelX, Screen.TwipsPerPixelY, UserControl.Width - Screen.TwipsPerPixelX, UserControl.Height - Screen.TwipsPerPixelY
    InputCover.Move 0, 0, UserControl.Width, UserControl.Height
    MyLabel.Move UserControl.Width / 2 - MyLabel.Width / 2, UserControl.Height * 0.8 - MyLabel.Height / 2
    MyIcon.Move UserControl.Width / 2 - 780 / 2, UserControl.Height * 0.8 / 2 - 780 / 2, 780, 780
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Image", imgData, StrConv("", vbFromUnicode))
    Call PropBag.WriteProperty("Content", MyLabel.Caption, "")
End Sub
