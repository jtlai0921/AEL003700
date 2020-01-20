VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "計算購書價款"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   4155
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
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "計算"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   480
      TabIndex        =   6
      Top             =   2520
      Width           =   1170
   End
   Begin VB.TextBox txtPrice 
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
      Left            =   2040
      TabIndex        =   3
      Text            =   " "
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtQty 
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
      Left            =   2040
      TabIndex        =   1
      Text            =   " "
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblResult 
      Caption         =   " "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblAmount 
      Caption         =   "總  價  款:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label lblPrice 
      Caption         =   "定        價:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblQty1 
      Caption         =   "購書數量:"
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
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalculate_Click()
   Dim qty As Integer, price As Integer, money As Integer
   price = Val(txtPrice): qty = Val(txtQty)
   If qty < 10 Then money = price * qty
   If qty >= 10 And qty < 50 Then money = price * qty * 0.9
   If qty >= 50 Then money = price * qty * 0.8
   lblResult = money
End Sub

Private Sub cmdEnd_Click()
   End
End Sub

