VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "顯示結果"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3735
   LinkTopic       =   "Form2"
   ScaleHeight     =   2445
   ScaleWidth      =   3735
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
      Caption         =   "回輸入畫面"
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
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  '單線固定
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Form2.Hide        '隱藏表單二
   Form1.Show        '顯示表單一
   Form1.Text1 = ""  '清除表單一文字方塊的內容
   Form1.Text1.SetFocus  '將輸入點設定在表單一的文字方塊
End Sub

