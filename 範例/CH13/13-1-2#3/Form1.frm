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
  Dim x As Integer, y As Integer
  x = InputBox("�п�J�Ĥ@�Ӿ��")
  y = InputBox("�п�J�ĤG�Ӿ��")
  If x / y = Int(x / y) Then
     Print "�i�H�㰣,���G�O:"; x / y
  Else
     Print "����㰣,���G�O:"; x / y
  End If
End Sub

