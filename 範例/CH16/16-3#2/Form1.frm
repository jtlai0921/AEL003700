VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  '��u�T�w
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
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.Label lblDisplay 
      BorderStyle     =   1  '��u�T�w
      BeginProperty Font 
         Name            =   "�з���"
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
      Caption         =   "�\��(&F)"
      Begin VB.Menu mnuDate 
         Caption         =   "���Ѥ��(&D)"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuTime 
         Caption         =   "�{�b�ɶ�(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "����(&X)"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "����(&A)"
      Begin VB.Menu mnuAuthor 
         Caption         =   "�@��(&U)"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuProgram 
         Caption         =   "���{��(&P)"
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
   lblDisplay = "�\�y�ڥ��ͭ���a���u�{���q��T�B�B���A"
   lblDisplay = lblDisplay + "�ثe��ܥ_�x��޾ǰ|��T�޲z�t����"
End Sub

Private Sub mnuDate_Click()
   lblDisplay = "���Ѥ���O" & Date
End Sub

Private Sub mnuEnd_Click()
   End
End Sub

Private Sub mnuProgram_Click()
   lblDisplay = "���{���O�]�w�\���P�]�p�U�������\�઺�{���X���d��"
End Sub

Private Sub mnuTime_Click()
   lblDisplay = "�{�b�ɶ��O" & Time
End Sub
