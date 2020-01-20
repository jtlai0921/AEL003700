VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "輸入資料"
   ClientHeight    =   2685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   ScaleHeight     =   2685
   ScaleWidth      =   4080
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "顯示結果"
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "請輸入姓名："
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Form2.Show                             '顯示表單二
   Form2.Label1 = "HELLO!" & Form1.Text1  '在表單二之標籤顯示結果
   Form1.Hide                             '隱藏表單ㄧ
End Sub

