VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   3480
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   3480
   StartUpPosition =   3  '�t�ιw�]��
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
