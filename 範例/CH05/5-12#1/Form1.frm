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
   Dim a As String, b As String, c As String
   Dim boy As String, girl As String, join As String
   a = "Visual"
   b = "BASIC"
   c = a + " " + b + "很好用"
   Print c
   boy = "羅蜜歐": girl = "朱麗葉"
   join = boy & "與" & girl
   Print join
   Print boy + "與" + girl
End Sub

