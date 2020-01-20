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
   StartUpPosition =   3  '系統預設值
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
