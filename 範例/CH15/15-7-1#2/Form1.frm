VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
   Open "a:\score.dat" For Output As #1
   nam = "張三":   chin = 78: eng = 82
   Write #1, nam, chin, eng
   nam = "李四":   chin = 66: eng = 76.5
   Write #1, nam, chin, eng
   nam = "王五":   chin = 82.5: eng = 90
   Write #1, nam, chin, eng
   Close #1
End Sub

