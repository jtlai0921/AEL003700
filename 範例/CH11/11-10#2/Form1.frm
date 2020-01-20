VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2850
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   2850
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "結束"
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   1200
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   1200
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   1920
      Picture         =   "Form1.frx":0000
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   1200
      Picture         =   "Form1.frx":0442
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   480
      Picture         =   "Form1.frx":0884
      Top             =   360
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim idx As Integer  '此敘述要安排在(一般)(宣告)中

Private Sub Command1_Click()
   End
End Sub

Private Sub Form_Load()
   For i = 0 To 2
      Image1(i).Left = 1200  '設定影像方塊的左邊界
   Next i
End Sub

Private Sub Timer1_Timer()
   idx = idx + 1
   If idx > 2 Then idx = 0
   Select Case idx
     Case 0
       Image1(0).Visible = True   '設定黃燈為可見
       Image1(1).Visible = False
       Image1(2).Visible = False
     Case 1
       Image1(0).Visible = False
       Image1(1).Visible = True   '設定紅燈為可見
       Image1(2).Visible = False
     Case 2
       Image1(0).Visible = False
       Image1(1).Visible = False
       Image1(2).Visible = True   '設定綠燈為可見
   End Select
End Sub
