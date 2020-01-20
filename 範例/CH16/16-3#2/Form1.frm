VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  '單線固定
   Caption         =   "Form1"
   ClientHeight    =   1755
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   2985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   2985
   StartUpPosition =   3  '系統預設值
   Begin VB.Label lblDisplay 
      BorderStyle     =   1  '單線固定
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2295
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuFunction 
      Caption         =   "功能(&F)"
      Begin VB.Menu mnuDate 
         Caption         =   "今天日期(&D)"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuTime 
         Caption         =   "現在時間(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "結束(&X)"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "關於(&A)"
      Begin VB.Menu mnuAuthor 
         Caption         =   "作者(&U)"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuProgram 
         Caption         =   "本程式(&P)"
         Shortcut        =   ^P
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuAuthor_Click()
   lblDisplay = "許慶芳先生原任榮民工程公司資訊處處長，"
   lblDisplay = lblDisplay + "目前轉至北台科技學院資訊管理系任教"
End Sub

Private Sub mnuDate_Click()
   lblDisplay = "今天日期是" & Date
End Sub

Private Sub mnuEnd_Click()
   End
End Sub

Private Sub mnuProgram_Click()
   lblDisplay = "此程式是設定功能表與設計各項對應功能的程式碼之範例"
End Sub

Private Sub mnuTime_Click()
   lblDisplay = "現在時間是" & Time
End Sub
