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
   Dim p As Single, r As Single, t As Single  'p������,r���~�Q�v,t�����Q�M
   Dim y As Integer       'y���~��
   p = InputBox("�п�J����", "��J����")
   r = InputBox("�п�J�~�Q�v", "��J�~�Q�v")
   y = 1
   Print "�~��", "���Q�M"
   Print                 '�Ť@�C
   Do While t < 2 * p '��t<2*p�N�~�����j��,�_�h����
      t = p * (1 + r / 100) ^ y  '�p���y�~�����Q�M
      Print y, t                 '��X�~�ƻP���Q�M
      y = y + 1                  '�֥[�@�~
   Loop
End Sub

