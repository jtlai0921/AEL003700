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
   password = InputBox$("�п�J�K�X", "�K�X�ˬd")
   If password = "1234" Then
      MsgBox "�q�L�K�X�ˬd�F!", vbOKOnly + vbExclamation, _
             "����!"
   Else
      feedback = MsgBox("�K�X���~!", vbYesNoCancel _
                 + vbCritical, "��p!")
      Select Case feedback
         Case vbYes
            MsgBox "�A���F�u�O(Y)�v�s", 0, "�O"
         Case vbNo
            MsgBox "�A���F�u�_(N)�v�s", 0, "�_"
         Case vbCancel
            MsgBox "�A���F�u�����v�s", 0, "����"
         Case Else
            MsgBox "�����������p"
      End Select
   End If
End Sub

