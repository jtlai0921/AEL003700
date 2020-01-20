VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   4230
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "切換"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Index           =   1
      Left            =   2400
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Index           =   0
      Left            =   480
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1515
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   If Picture1(0).Visible = False Then    '假如左邊圖看不見就…
     Picture1(0).Picture = Picture1(1).Picture  '將右邊圖設定給左邊方塊
     Picture1(0).Visible = True           '設定左邊圖看得見
     Picture1(1).Visible = False          '設定右邊圖看不見
   Else                            '左邊圖看得見的情況
     Picture1(1).Picture = Picture1(0).Picture  '將左邊圖設定給右邊方塊
     Picture1(1).Visible = True           '設定右邊圖看得見
     Picture1(0).Visible = False          '設定左邊圖看不見
  End If
End Sub
