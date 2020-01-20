VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
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
      Left            =   2160
      TabIndex        =   7
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定"
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
      Left            =   2160
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.CheckBox Check4 
      Caption         =   "聊天"
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
      Left            =   360
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      Caption         =   "音樂"
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
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   15
      Left            =   720
      TabIndex        =   3
      Top             =   1080
      Width           =   135
   End
   Begin VB.CheckBox Check1 
      Caption         =   "運動"
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
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.CheckBox Check2 
      Caption         =   "游泳"
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
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
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
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "請勾選自己有興趣的活動:"
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
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Dim interest As String  '設定字串變數,用來連結各項興趣
   Dim n As Integer        '設定n為整數,用來累計有幾項興趣
   If Check1.Value = 1 Then _
      interest = interest + " 運動": n = n + 1
   If Check2.Value = 1 Then _
      interest = interest + " 游泳": n = n + 1
   If Check3.Value = 1 Then _
      interest = interest + " 音樂": n = n + 1
   If Check4.Value = 1 Then _
      interest = interest + " 聊天": n = n + 1
   Label2.Caption = "你的興趣有:" + interest + _
                    Str(n) + "項"
End Sub

Private Sub Command2_Click()
   End
End Sub

