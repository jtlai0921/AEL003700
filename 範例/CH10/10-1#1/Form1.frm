VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   3750
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
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
      Height          =   372
      Left            =   1320
      TabIndex        =   1
      Top             =   1440
      Width           =   972
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   372
      LargeChange     =   10
      Left            =   600
      Max             =   100
      TabIndex        =   0
      Top             =   360
      Width           =   2532
   End
   Begin VB.Label Label2 
      Caption         =   "0"
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
      Left            =   2640
      TabIndex        =   3
      Top             =   960
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "目前捲動鈕位置:"
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
      Left            =   720
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   End
End Sub

Private Sub HScroll1_Change()
   Label2 = HScroll1.Value
End Sub

Private Sub HScroll1_Scroll()
   Label2 = HScroll1.Value
End Sub
