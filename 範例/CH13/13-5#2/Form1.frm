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
  Dim x As Integer, y As Integer, sum As Integer, score As Integer
  Randomize
  For i = 1 To 10
    x = Int(Rnd * 90) + 10       '����10~99�������H�N���
    y = Int(Rnd * 90) + 10       '����10~99�������H�N���
    sum = InputBox("�п�J" + Str(x) + "+" + Str(y) + "=?")
    If sum = x + y Then          '���諸���p
      MsgBox "����!�A�o�D����F!"
      score = score + 10
    Else                         '���������p
      MsgBox "��p!�A�o�D�����F! ���T�����׬O:" + Str(x + y)
    End If
  Next i
  MsgBox "���絲��!�A���o���O:" + Str(score)
End Sub

