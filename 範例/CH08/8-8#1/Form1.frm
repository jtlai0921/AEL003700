VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "電子時鐘"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   4245
   StartUpPosition =   3  '系統預設值
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   240
   End
   Begin VB.Label Label2 
      Caption         =   " "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "現在時間 "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   20.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   Timer1.Interval = 1000  '設定時間間隔為1秒
End Sub

Private Sub Timer1_Timer() '每隔1秒就驅動此程序
   Label2 = Time           '顯示目前時間
End Sub
