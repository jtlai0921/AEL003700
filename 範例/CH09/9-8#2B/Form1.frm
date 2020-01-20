VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "儲蓄是美德"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
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
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   2640
      Width           =   1290
   End
   Begin VB.CommandButton cmdCompute 
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
      Left            =   600
      TabIndex        =   11
      Top             =   2640
      Width           =   1290
   End
   Begin VB.TextBox txtDays 
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
      TabIndex        =   10
      Text            =   " "
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtDeposit 
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
      TabIndex        =   9
      Text            =   " "
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label9 
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
      Left            =   4320
      TabIndex        =   8
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label lblTotalDeposit 
      Alignment       =   1  '靠右對齊
      BorderStyle     =   1  '單線固定
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
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "全部存款總額"
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
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label lblLastDeposit 
      Alignment       =   1  '靠右對齊
      BorderStyle     =   1  '單線固定
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
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "最後一天應存金額"
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
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label8 
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
      Left            =   4320
      TabIndex        =   3
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label7 
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
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "持續存款天數"
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
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "第一天存款額"
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
      Left            =   240
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
Private Sub cmdCompute_Click()
  Dim day_deposit As Double, total_deposit As Double  '當天存款與累積存款
  Dim days As Integer                       '持續存款天數
  deposit = Val(txtDeposit)
  days = Val(txtDays)
  total_deposit = deposit       '第一天的累積存款等於當天存款
  For i = 2 To days
    day_deposit = deposit * 2 ^ (i - 1) '每天存款為第一天存款的2^(i-1)倍
    total_deposit = total_deposit + day_deposit  '累積總存款金額
  Next i
  Rem 顯示結果
  lblLastDeposit.Caption = Format(day_deposit, "#########,###")
  lblTotalDeposit.Caption = Format(total_deposit, "#########,###")
End Sub

Private Sub cmdEnd_Click()
   End
End Sub
