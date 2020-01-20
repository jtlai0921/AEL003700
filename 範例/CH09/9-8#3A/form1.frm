VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "收銀機"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   4260
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
      Height          =   420
      Left            =   1440
      TabIndex        =   8
      Top             =   1800
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
      Height          =   420
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.TextBox txtCash 
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
      Left            =   1440
      TabIndex        =   5
      Text            =   " "
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txtAmount 
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
      Left            =   1440
      TabIndex        =   4
      Text            =   " "
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "找錢"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   1455
      Begin VB.Label Label22 
         Caption         =   "個"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   26
         Top             =   2040
         Width           =   255
      End
      Begin VB.Label Label21 
         Caption         =   "個"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   25
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label20 
         Caption         =   "個"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   24
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label19 
         Caption         =   "個"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   23
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label18 
         Caption         =   "個"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   22
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label17 
         Caption         =   "個"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lbl1 
         BorderStyle     =   1  '單線固定
         Caption         =   " "
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   20
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label lbl5 
         BorderStyle     =   1  '單線固定
         Caption         =   " "
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lbl10 
         BorderStyle     =   1  '單線固定
         Caption         =   " "
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   18
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label lbl50 
         BorderStyle     =   1  '單線固定
         Caption         =   " "
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   17
         Top             =   960
         Width           =   375
      End
      Begin VB.Label lbl100 
         BorderStyle     =   1  '單線固定
         Caption         =   " "
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   16
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lbl500 
         BorderStyle     =   1  '單線固定
         Caption         =   " "
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "1元"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "5元"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "10元"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "50元"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "100元"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "500元"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label lblChange 
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
      Left            =   1440
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "應找金額"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "收款金額"
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
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "購物金額"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCompute_Click()
   Dim amount As Integer, cash As Integer, change As Integer, _
       piece As Integer, unit As Integer
   amount = Val(txtAmount)
   cash = Val(txtCash)
   change = cash - amount
   If change < 0 Then    '假如應找回的金額為負值,就顯示錯誤訊息, 並結束執行
      MsgBox "收款金額不應低於購物金額,請重新執行程式", 48, "輸入資料錯誤!"
      End
   End If
   lblChange = change            '顯示應找回的金額
   For i = 1 To 6
      unit = Choose(i, 500, 100, 50, 10, 5, 1)
      piece = 0
      Do While change >= unit
         piece = piece + 1
         change = change - unit
      Loop
      Select Case i
         Case 1: lbl500 = piece  '顯示找500元錢幣個數
         Case 2: lbl100 = piece  '顯示找100元錢幣個數
         Case 3: lbl50 = piece   '顯示找50元錢幣個數
         Case 4: lbl10 = piece   '顯示找10元錢幣個數
         Case 5: lbl5 = piece    '顯示找5元錢幣個數
         Case 6: lbl1 = piece    '顯示找1元錢幣個數
      End Select
   Next i
End Sub

Private Sub cmdEnd_Click()
   End
End Sub
