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
   StartUpPosition =   3  '╰参箇砞
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   Dim n As Integer, pass As Integer  'n厩ネ计,passの计
   Dim score As Single     'scoreだ计
   
   n = InputBox("叫块厩ネ计", "块厩ネ计")
   For i = 1 To n
      score = InputBox("叫块だ计", "块だ计")
      If score >= 60 Then pass = pass + 1
  Next i
  Print "の计  ="; pass; ""
  Print "ぃの计="; n - pass; ""
End Sub

