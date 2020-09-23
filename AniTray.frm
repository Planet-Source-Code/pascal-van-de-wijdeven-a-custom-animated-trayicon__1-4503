VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3270
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4740
   Icon            =   "AniTray.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   120
      Top             =   240
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin VB.Menu Popup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Counter As Integer
Public IconObject As Object

Private Sub Exit_Click()
    Unload Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    delIcon IconObject.Handle
    delIcon Form1.Icon.Handle
End Sub


Private Sub Form_Load()
    Set IconObject = Form1.Icon
    AddIcon Form1, IconObject.Handle, IconObject, "Animated TrayIcon"
End Sub

Private Sub Timer1_Timer()
    Counter = Counter + 1
    Form1.Icon = ImageList1.ListImages(Counter).Picture
    If Counter > ImageList1.ListImages.Count - 1 Then Counter = 0
    modIcon Form1, IconObject.Handle, Form1.Icon, "Animated TrayIcon"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static Message As Long
    Message = X / Screen.TwipsPerPixelX
    Select Case Message
    Case WM_RBUTTONUP:
        Me.PopupMenu Popup
    End Select
End Sub
