VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   4155
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢"
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
      Left            =   2040
      TabIndex        =   10
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "新增"
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
      Left            =   1080
      TabIndex        =   9
      Top             =   2400
      Width           =   855
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
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "開檔"
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
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox txtEng 
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
      Text            =   " "
      Top             =   1800
      Width           =   735
   End
   Begin VB.TextBox txtChin 
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
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox txtName 
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
      TabIndex        =   4
      Text            =   " "
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "英文成績:"
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
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "國文成績:"
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
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "姓        名:"
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
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '單線固定
      Caption         =   " 學 生 成 績 登 錄 作 業"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.75
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
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOpen_Click()
   Open "a:\score.dat" For Output As #1
   Close #1
End Sub

Private Sub cmdAdd_Click()
   Open "a:\score.dat" For Append As #1
   Write #1, txtName, txtChin, txtEng
   Close #1
   txtName = "": txtChin = "": txtEng = ""
End Sub


Private Sub cmdQuery_Click()
   Open "a:\score.dat" For Input As #1
   Dim find As Boolean
   find = False
   Do While Not EOF(1)
      Input #1, nam, chin, eng
      If nam = txtName Then
         find = True
         txtChin = chin
         txtEng = eng
         Exit Do
     End If
   Loop
   If find Then
      MsgBox "有找到,並已顯示!"
   Else
      MsgBox "檔案中沒有要找的姓名:" + txtName
   End If
   Close #1
   txtName = "": txtChin = "": txtEng = ""
End Sub

Private Sub cmdEnd_Click()
   End
End Sub


