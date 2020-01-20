VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "加法器"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2100
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   2100
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Text2 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Text            =   " "
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  '靠右對齊
      BorderStyle     =   1  '單線固定
      Caption         =   " "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   14.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   1800
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label1 
      Caption         =   " ＋"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()  '按一下Label1,就執行此程序
   Label2 = Val(Text1) + Val(Text2)
End Sub
