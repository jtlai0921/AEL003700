VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "計算電費"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3690
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   3690
   StartUpPosition =   3  '系統預設值
   Begin VB.Frame Frame1 
      Caption         =   "用電種類"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   480
      TabIndex        =   7
      Top             =   840
      Width           =   2655
      Begin VB.OptionButton optBusiness 
         Caption         =   "營業用"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optHome 
         Caption         =   "家庭用"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
   End
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
      TabIndex        =   4
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdCal 
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
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtDegree 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1680
      TabIndex        =   1
      Text            =   " "
      Top             =   255
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "元"
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
      Left            =   2520
      TabIndex        =   6
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label LblFee 
      Alignment       =   1  '靠右對齊
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
      Left            =   1440
      TabIndex        =   5
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   " 電費:"
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
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "用電度數:"
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
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   optHome.Value = True  '預設選用了家庭用電
End Sub
Private Sub cmdCal_Click()
   Dim fee As Single, degree As Single
   degree = Val(txtDegree)
   If optHome.Value = True Then  '假如選家庭用電就
      Call home(degree, fee)     '呼叫副程式home
   Else                          '否則
      Call business(degree, fee) '呼叫副程式Business
   End If
   LblFee = fee      '將電費放入顯示結果的位置
End Sub

Private Sub cmdEnd_Click()
   End
End Sub

Private Sub home(d As Single, f As Single)
   Select Case d
      Case Is <= 100
         f = 2.4 * d
      Case Is <= 300
         f = 2.4 * 100 + 3.1 * (d - 100)
      Case Else
         f = 2.4 * 100 + 3.1 * 200 + 4.1 * (d - 300)
   End Select
End Sub
Private Sub business(d As Single, f As Single)
   If d <= 300 Then
      f = 5.9 * d
   Else
      f = 5.9 * 300 + 6.7 * (d - 300)
   End If
End Sub
