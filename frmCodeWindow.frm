VERSION 5.00
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "Codejock.SyntaxEdit.v15.3.1.ocx"
Begin VB.Form frmCodeWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "代码窗口"
   ClientHeight    =   5175
   ClientLeft      =   3540
   ClientTop       =   3060
   ClientWidth     =   8865
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmCodeWindow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   8865
   Begin XtremeSyntaxEdit.SyntaxEdit SyntaxEdit 
      Height          =   1935
      Left            =   1080
      TabIndex        =   0
      Top             =   1200
      Width           =   3015
      _Version        =   983043
      _ExtentX        =   5318
      _ExtentY        =   3413
      _StockProps     =   84
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Consolas"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EnableSyntaxColorization=   -1  'True
      ShowLineNumbers =   -1  'True
      ShowSelectionMargin=   0   'False
      ShowScrollBarVert=   -1  'True
      ShowScrollBarHorz=   -1  'True
      EnableVirtualSpace=   0   'False
      EnableAutoIndent=   -1  'True
      ShowWhiteSpace  =   0   'False
      ShowCollapsibleNodes=   -1  'True
      AutoCompleteWndWidth=   160
      EnableEditAccelerators=   -1  'True
   End
   Begin VB.PictureBox picSelMargin 
      Appearance      =   0  'Flat
      BackColor       =   &H00333333&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   240
      ScaleHeight     =   1935
      ScaleWidth      =   255
      TabIndex        =   4
      Top             =   1320
      Width           =   255
   End
   Begin DragControlsIDE.DarkComboBox comObject 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   556
      Items0          =   ""
      ITEM_COUNT      =   0
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin DragControlsIDE.DarkWindowBorder DarkWindowBorder 
      Left            =   7560
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      Thickness       =   4
      MinWidth        =   150
      MinHeight       =   100
   End
   Begin DragControlsIDE.DarkTitleBar DarkTitleBar 
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "代码窗口"
      MaxButtonVisible=   0   'False
      MinButtonVisible=   0   'False
      BindCaption     =   -1  'True
      Picture         =   "frmCodeWindow.frx":1BCC2
   End
   Begin DragControlsIDE.DarkWindowBorder DarkWindowBorderSizer 
      Left            =   8280
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      Thickness       =   3
      FocusedColor    =   3157293
      NotFocusedColor =   3157293
      MinWidth        =   150
      MinHeight       =   100
   End
   Begin DragControlsIDE.DarkComboBox comEvent 
      Height          =   315
      Left            =   4560
      TabIndex        =   2
      Top             =   660
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   556
      Items0          =   ""
      ITEM_COUNT      =   0
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft YaHei UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCodeWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WindowObj    As Object                                                                                           '窗口自身
Public FileIndex    As Long                                                                                             '在CurrentProject.Files对应的文件序号

Private Sub DarkTitleBar_GotFocus()
    On Error Resume Next
    
    Me.SyntaxEdit.SetFocus
End Sub

Private Sub Form_Load()
    Me.Caption = Lang_CodeWindow_Caption
    Me.DarkTitleBar.MaxButtonVisible = True
    Me.DarkTitleBar.MinButtonVisible = True
    
    '设置代码框属性
    Me.DarkTitleBar.Top = Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelY
    Me.SyntaxEdit.Move Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelX + Me.picSelMargin.Width, _
        Me.DarkTitleBar.Height + Me.comObject.Height + 240 + Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelY
    Me.picSelMargin.Move Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelX, Me.SyntaxEdit.Top, 300, Me.SyntaxEdit.Height
    Me.SyntaxEdit.PaintManager.BackColor = RGB(28, 28, 28)
    Me.SyntaxEdit.PaintManager.LineNumberBackColor = RGB(28, 28, 28)
    Me.SyntaxEdit.PaintManager.LineNumberTextColor = RGB(86, 156, 214)
    Me.SyntaxEdit.DataManager.FileExt = ".cpp"
    Me.SyntaxEdit.ConfigFile = App.Path & "\SyntaxEdit.ini"
    
    '设置窗口子类化，处理最大化问题及处理任务栏右键关闭
    Dim lpObj               As Long                                                                                     '指向窗口自身的物件指针
    Set WindowObj = Me
    lpObj = ObjPtr(WindowObj)                                                                                           '获取指向窗口自身的物件指针
    SetPropA Me.hWnd, "WindowObj", lpObj                                                                                '记录窗口的物件地址，供子类化卸载窗体用
    'SetPropA Me.hWnd, "PrevWndProc", SetWindowLongA(Me.hWnd, GWL_WNDPROC, AddressOf MainWindowMaximizeCloseFixProc)    '[ToDo]
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If IsExiting Then
        '恢复窗口子类化
        SetWindowLongA Me.hWnd, GWL_WNDPROC, GetPropA(Me.hWnd, "PrevWndProc")
    Else
        Cancel = 1
        Me.Hide
    End If
End Sub

Public Sub Form_Resize()
    On Error Resume Next
    
    '根据标题栏是否显示来调整控件位置
    If Me.DarkTitleBar.Visible = True Then
        Me.comObject.Top = Me.DarkTitleBar.Height + 165
        Me.comEvent.Top = Me.comObject.Top
        Me.SyntaxEdit.Top = Me.comEvent.Top + Me.comEvent.Height + 240
    Else
        Me.comObject.Top = 120
        Me.comEvent.Top = 120
        Me.SyntaxEdit.Top = 120 + Me.comObject.Height + 120
    End If
    Me.picSelMargin.Top = Me.SyntaxEdit.Top
    
    '设置代码框大小
    Me.SyntaxEdit.Width = Me.ScaleWidth - Me.SyntaxEdit.Left - Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelX
    Me.SyntaxEdit.Height = Me.ScaleHeight - Me.SyntaxEdit.Top - Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelY
    Me.picSelMargin.Height = Me.SyntaxEdit.Height
    
    '设置组合框大小和位置
    Me.comObject.Width = (Me.ScaleWidth - 480) / 2
    Me.comEvent.Width = Me.comObject.Width
    Me.comEvent.Left = Me.comObject.Width + 360
End Sub

Private Sub SyntaxEdit_TextChanged(ByVal nRowFrom As Long, ByVal nRowTo As Long, ByVal nActions As Long)
    CurrentProject.Files(FileIndex).Changed = True                                                      '代码框的内容一旦更改，就把文件视为更改了
    
End Sub

