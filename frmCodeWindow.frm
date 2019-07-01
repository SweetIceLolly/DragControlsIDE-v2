VERSION 5.00
Object = "{CA73588D-282F-4592-9369-A61CC244FADA}#15.3#0"; "SYNTAX~1.OCX"
Begin VB.Form frmCodeWindow 
   Appearance      =   0  'Flat
   BackColor       =   &H00302D2D&
   BorderStyle     =   0  'None
   Caption         =   "代码窗口"
   ClientHeight    =   5172
   ClientLeft      =   3540
   ClientTop       =   3060
   ClientWidth     =   8868
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmCodeWindow.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5172
   ScaleWidth      =   8868
   Begin XtremeSyntaxEdit.SyntaxEdit SyntaxEdit 
      Height          =   1935
      Left            =   240
      TabIndex        =   3
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
      ShowSelectionMargin=   -1  'True
      ShowScrollBarVert=   -1  'True
      ShowScrollBarHorz=   -1  'True
      EnableVirtualSpace=   0   'False
      EnableAutoIndent=   -1  'True
      ShowWhiteSpace  =   0   'False
      ShowCollapsibleNodes=   -1  'True
      AutoCompleteWndWidth=   160
      EnableEditAccelerators=   -1  'True
   End
   Begin DragControlsIDE.DarkComboBox comObject 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   660
      Width           =   4095
      _ExtentX        =   7218
      _ExtentY        =   550
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
      _ExtentX        =   677
      _ExtentY        =   677
      Thickness       =   4
      MinWidth        =   150
      MinHeight       =   100
   End
   Begin DragControlsIDE.DarkTitleBar DarkTitleBar 
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8865
      _ExtentX        =   15642
      _ExtentY        =   868
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
      BindCaption     =   -1  'True
      Picture         =   "frmCodeWindow.frx":1BCC2
   End
   Begin DragControlsIDE.DarkWindowBorder DarkWindowBorderSizer 
      Left            =   8280
      Top             =   4560
      _ExtentX        =   677
      _ExtentY        =   677
      Thickness       =   3
      FocusedColor    =   3157293
      NotFocusedColor =   3157293
      MinWidth        =   150
      MinHeight       =   100
   End
   Begin DragControlsIDE.DarkComboBox comEvent 
      Height          =   315
      Left            =   4560
      TabIndex        =   1
      Top             =   660
      Width           =   4095
      _ExtentX        =   7218
      _ExtentY        =   550
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

Private Sub DarkTitleBar_GotFocus()
    On Error Resume Next
    
    Me.SyntaxEdit.SetFocus
End Sub

Private Sub Form_Load()
    '设置代码框属性
    Me.DarkTitleBar.Top = Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelY
    Me.SyntaxEdit.Move Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelX, _
        Me.DarkTitleBar.Height + Me.comObject.Height + 240 + Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelY
    Me.SyntaxEdit.PaintManager.BackColor = RGB(28, 28, 28)
    Me.SyntaxEdit.PaintManager.LineNumberBackColor = RGB(28, 28, 28)
    Me.SyntaxEdit.PaintManager.LineNumberTextColor = RGB(86, 156, 214)
    Me.SyntaxEdit.DataManager.FileExt = ".cpp"
    Me.SyntaxEdit.ConfigFile = App.Path & "\SyntaxEdit.ini"
    
    '设置窗口子类化，处理最大化问题及处理任务栏右键关闭
    '设置窗口子类化，处理最大化问题及处理任务栏右键关闭
    Dim lpObj               As Long                                                                                     '指向窗口自身的物件指针
    Set WindowObj = Me
    lpObj = ObjPtr(Me)                                                                                                  '获取指向窗口自身的物件指针
    SetPropA Me.hWnd, "WindowObj", lpObj                                                                                '记录窗口的物件地址，供子类化卸载窗体用
    SetPropA Me.hWnd, "PrevWndProc", SetWindowLongA(Me.hWnd, GWL_WNDPROC, AddressOf MainWindowMaximizeCloseFixProc)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '恢复窗口子类化
    SetWindowLongA Me.hWnd, GWL_WNDPROC, GetPropA(Me.hWnd, "PrevWndProc")
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
    
    '设置代码框大小
    Me.SyntaxEdit.Width = Me.ScaleWidth - Me.SyntaxEdit.Left - Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelX
    Me.SyntaxEdit.Height = Me.ScaleHeight - Me.SyntaxEdit.Top - Me.DarkWindowBorderSizer.Thickness * Screen.TwipsPerPixelY
    
    '设置组合框大小和位置
    Me.comObject.Width = (Me.ScaleWidth - 480) / 2
    Me.comEvent.Width = Me.comObject.Width
    Me.comEvent.Left = Me.comObject.Width + 360
End Sub
