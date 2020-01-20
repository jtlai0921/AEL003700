VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "分數轉換等級"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdEnd 
      Caption         =   "結束"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "轉換"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtScore 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Text            =   " "
      Top             =   360
      Width           =   975
   End
   Begin VB.Label lblResult 
      Alignment       =   2  '置中對齊
      BorderStyle     =   1  '單線固定
      Caption         =   " "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   24
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2160
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lblGrade 
      Caption         =   "等    級  →"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblScore 
      Caption         =   "請輸入分數:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConvert_Click()
   Dim s As Single     '設定score為單精數變數
   s = Val(txtScore)   '將輸入文字方塊的資料轉成數值
   If s >= 80 Then
      lblResult = "A"
   ElseIf s >= 70 Then
      lblResult = "B"
   ElseIf s >= 60 Then
      lblResult = "C"
   Else
      lblResult = "D"
   End If
End Sub

Private Sub cmdEnd_Click()
   End
End Sub
