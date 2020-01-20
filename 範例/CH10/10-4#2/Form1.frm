VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "清單方塊的基本操作"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4560
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
      Left            =   3240
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "全清"
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
      Left            =   3240
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "刪除"
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
      Left            =   3240
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.ListBox lstData 
      Height          =   1320
      ItemData        =   "Form1.frx":0000
      Left            =   240
      List            =   "Form1.frx":0002
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton cmdInput 
      Caption         =   "輸入"
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
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtInput 
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
      Text            =   " "
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "已輸入的資料項:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   2292
   End
   Begin VB.Label Label1 
      Caption         =   "請輸入要新增的資料:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2292
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdInput_Click()
   newdata = txtInput.Text   '將文字方塊的內容存進變數newdata中
   For i = 0 To lstData.ListCount - 1
      If lstData.Selected(i) Then   '假如第i項被選取就
         lstData.AddItem newdata, i '將變數newdata的內容加入清單中
         flag = 1    '設定有選取資料項的狀況
         Exit For    '跳離迴圈
      End If
   Next i
   If flag = 0 Then lstData.AddItem newdata '沒有選取資料項就加在尾巴
   txtInput.Text = ""      '將文字方塊清成空字串
   txtInput.SetFocus       '將駐點設定在輸入資料的文字方塊上
End Sub

Private Sub cmdDelete_Click()
   For i = 0 To lstData.ListCount - 1
      If lstData.Selected(i) Then   '假如第i項被選取就
         lstData.RemoveItem i       '刪除第i項
         Exit For                    '跳離迴圈
      End If
   Next i
End Sub

Private Sub cmdClear_Click()
   lstData.Clear        '將清單方塊的內容全部清除掉
End Sub

Private Sub cmdEnd_Click()
   End
End Sub

