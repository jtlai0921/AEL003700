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
   w = ScaleWidth  '���o���iø�ϳ������e��
   h = ScaleHeight '���o���iø�ϳ���������
   Do
      n = n + 1
      r = n * 100    '�C�骺�b�|�W�[100Twips
      '�b�|�W�L�e�שΰ��ת��@�b,�N�����j��
      If r > w / 2 Or r > h / 2 Then Exit Do
      Circle (w / 2, h / 2), r, QBColor(n)  '�e��
  Loop
End Sub

