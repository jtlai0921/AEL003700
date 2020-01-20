VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "求最大公因數與最小公倍數"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4020
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   4020
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
      Left            =   1920
      TabIndex        =   9
      Top             =   2520
      Width           =   975
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
      Height          =   375
      Left            =   600
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtno2 
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
      Left            =   2640
      TabIndex        =   3
      Text            =   " "
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtno1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2640
      TabIndex        =   2
      Text            =   " "
      Top             =   225
      Width           =   615
   End
   Begin VB.Label lblLCM 
      BorderStyle     =   1  '單線固定
      Caption         =   " "
      DataField       =   " "
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
      Left            =   2640
      TabIndex        =   7
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblGCD 
      BorderStyle     =   1  '單線固定
      Caption         =   " "
      DataField       =   " "
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
      Left            =   2640
      TabIndex        =   6
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "最小公倍數為:"
      DataField       =   " "
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
      Left            =   840
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "最大公因數為:"
      DataField       =   " "
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
      Left            =   840
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "請輸入第二個整數:"
      DataField       =   " "
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
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "請輸入第一個整數:"
      DataField       =   " "
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCompute_Click()
  Do   '檢查第一個數是否為正整數,不是就要求重新輸入
    m = Val(txtno1)
    If m > 0 And m = Int(m) Then Exit Do
    MsgBox "第一個數非正整數,請重新輸入後再按計算鈕"
    txtno1 = ""           '將輸入的第一個數清成空白
    txtno1.SetFocus  '設定駐點物件,即插入點停在此格
    Exit Sub         '跳離此程序
  Loop
  Do   '檢查第二個數是否為正整數,不是就要求重新輸入
    n = Val(txtno2)
    If n > 0 And n = Int(n) Then Exit Do
    MsgBox "第二個數非正整數,請重新輸入後再按計算鈕"
    txtno2 = ""
    txtno2.SetFocus
    Exit Sub
  Loop
  If m > n Then k = n Else k = m   '設定m與n的較小值
  Rem 找出兩數的最大公因數
  For i = k To 1 Step -1
    If m / i = Int(m / i) And n / i = Int(n / i) Then Exit For  '找到了
  Next i
  lblGCD = i                '顯示最大公因數
  lblLCM = m * n / i        '顯示最小公倍數
End Sub

Private Sub cmdEnd_Click()
   End
End Sub

