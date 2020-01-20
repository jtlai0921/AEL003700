VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "MaxLength=14"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "MaxLength=0"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Text1.MaxLength = 0
   Text1 = "Visual BASIC易學易用"
End Sub

Private Sub Command2_Click()
   Text1.MaxLength = 14
   Text1 = "Visual BASIC易學易用"
End Sub
