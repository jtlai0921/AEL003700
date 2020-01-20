VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "選取檔案"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   5940
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton Command1 
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
      Left            =   4560
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.FileListBox File1 
      Height          =   2610
      Left            =   3360
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.DirListBox Dir1 
      Height          =   1560
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label8 
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
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   3480
      Width           =   3615
   End
   Begin VB.Label Label5 
      Caption         =   "工作目錄:"
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
      TabIndex        =   11
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label7 
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
      Height          =   255
      Left            =   1800
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "工作磁碟機:"
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
      TabIndex        =   9
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "檔    案:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "目錄:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "磁碟機:"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label9 
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
      Height          =   255
      Left            =   1800
      TabIndex        =   5
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "選取的檔案:"
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
      Top             =   4200
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Drive1_Change()
   Dir1.Path = Drive1.Drive '將選定的磁碟機設定為目錄清單的路徑
   Label7 = Drive1.Drive    '顯示選定的磁碟機
   Label9 = ""              '將顯示選取的檔案名稱位置清成空字串
End Sub

Private Sub Dir1_Change()
   File1.Path = Dir1.Path   '將選定的目錄設定為檔案清單的路徑
   Label8 = Dir1.Path       '顯示選定的目錄
   Label9 = ""              '將顯示選取的檔案名稱位置清成空字串
End Sub


Private Sub File1_Click()
   Label9 = File1.FileName  '顯示選取的檔案名稱
End Sub


Private Sub Command1_Click()  '按「結束」鈕就執行此程序
   End
End Sub


