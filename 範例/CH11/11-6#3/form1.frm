VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5265
   StartUpPosition =   3  '�t�ιw�]��
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   Dim w As Integer
   Dim week As Variant  '�]�wa�����w���ܼ�
   Do
   w = InputBox("�п�J0~6���@�Ӿ��")
   If w >= 0 And w <= 6 Then Exit Do
   MsgBox "��J�ƭȶW�X�d��,�Э��s��J"
   Loop
   week = Array("��", "�@", "�G", "�T", "�|", "��", "��")
   Print "�P��"; week(w)
End Sub

