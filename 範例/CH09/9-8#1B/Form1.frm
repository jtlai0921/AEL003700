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
   Dim nopass As Integer, pass As Integer  'pass���ή�H��,nopass�����ή�H��
   Dim score As Single, total As Single     'score���ӤH����,total�����Z�`��
   Do
      score = InputBox("�п�J����", "��J����")
      If score < 0 Then Exit Do   '�p�G���Ƭ��t��,�N�����j��
      If score <= 100 Then        '���ƥ��`
         If score >= 60 Then pass = pass + 1 Else nopass = nopass + 1
         total = total + score    '�֥[���Z�`��
      Else                        '���ƶW�L100,�N��X���~�T��
         MsgBox "��J���~-���ƶW�L100,�Э��s��J", 48, "��ƿ��~"
      End If
   Loop
   Rem ��X���G
   Print "�ή�H��  ="; pass; "�H"
   Print "���ή�H��="; nopass; "�H"
   Print "���Z����  ="; total / (pass + nopass); "��"
End Sub

