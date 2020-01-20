VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4560
   StartUpPosition =   3  '系統預設值
   Begin VB.ListBox List1 
      Height          =   1500
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
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
   Begin VB.TextBox Text1 
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
      Caption         =   "已輸入的影片名稱如下:"
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
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "請輸入影片名稱:"
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
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   ItemData = Text1.Text   '將文字方塊的內容存進變數ItemData中
   List1.AddItem ItemData  '將變數ItemData的內容加入清單方塊中
   Text1.Text = ""         '將文字方塊清成空字串
   Text1.SetFocus          '將駐點設定在輸入資料的文字方塊上
End Sub
