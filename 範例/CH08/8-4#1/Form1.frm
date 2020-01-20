VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   12
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   4785
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command2 
      Caption         =   "結束"
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定"
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "學歷"
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   3495
      Begin VB.OptionButton Option7 
         Caption         =   "國中小"
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option6 
         Caption         =   "高中職"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton Option5 
         Caption         =   "大專"
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option4 
         Caption         =   "碩士"
         Height          =   255
         Left            =   1320
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "博士"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "性別"
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   3495
      Begin VB.OptionButton Option2 
         Caption         =   "女"
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "男"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Text            =   " "
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  '單線固定
      Caption         =   " "
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   3120
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "請輸入姓名:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Dim sex As String, education As String
   If Option1 Then sex = "先生" Else sex = "小姐"
   If Option3 Then
      education = "博士"
   ElseIf Option4 Then
      education = "碩士"
   ElseIf Option5 Then
      education = "大專"
   ElseIf Option6 Then
      education = "高中職"
   Else
      education = "國中小"
   End If
   Label2.Caption = "HI!" + Text1.Text + sex + _
          ",學歷:" + education + ",歡迎來學VB!"
End Sub

Private Sub Command2_Click()
   End
End Sub
