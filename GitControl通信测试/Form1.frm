VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   8895
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   7080
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Winsock.Bind 6028
    Winsock.Listen
End Sub


Private Sub Winsock_ConnectionRequest(ByVal requestID As Long)
    Winsock.Close
    Winsock.Accept requestID
End Sub

Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    Dim a As String
    Winsock.GetData a
    MsgBox a
End Sub


