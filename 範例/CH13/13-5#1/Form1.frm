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
  Dim s As Integer, g As Integer, t As Integer
  Randomize
  s = Int(Rnd * 50) + 1
    Do
    t = t + 1
    g = InputBox("�п�J�A�Ҳq���ƭ�(1~50)")
    Select Case g
      Case s
        MsgBox "����!�A�q��F! �`�@�q�F" + Str(t) + "��"
        Exit Do
      Case Is < s
        MsgBox "��" + Str(t) + "���q���ƭ�" + Str(g) + "�ӧC�F!�ЦA�q�@��!"
      Case Else
        MsgBox "��" + Str(t) + "���q���ƭ�" + Str(g) + "�Ӱ��F!�ЦA�q�@��!"
    End Select
  Loop
End Sub

