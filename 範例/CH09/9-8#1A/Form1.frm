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
   StartUpPosition =   3  '�t�ιw�]��
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   Dim n As Integer, pass As Integer  'n���ǥͤH��,pass���ή�H��
   Dim score As Single     'score������
   
   n = InputBox("�п�J�ǥͤH��", "��J�ǥͤH��")
   For i = 1 To n
      score = InputBox("�п�J����", "��J����")
      If score >= 60 Then pass = pass + 1
  Next i
  Print "�ή�H��  ="; pass; "�H"
  Print "���ή�H��="; n - pass; "�H"
End Sub

