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
  s = InputBox("�п�J����")
  Select Case s
     Case Is < 60
        Print "���ή�"
        Print "�Цh�Υ\�I"
     Case 60 To 89
        Print "�ή�"
        Print "���ߡI"
     Case Else
        Print "�u��"
        Print "�o�����@�i�I"
  End Select
End Sub

