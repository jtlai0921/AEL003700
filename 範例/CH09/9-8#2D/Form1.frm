VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "�x�W�O���w"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   4785
   StartUpPosition =   3  '�t�ιw�]��
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Activate()
  Dim day_deposit As Double, total_deposit As Double  '��Ѧs�ڻP�ֿn�s��
  Dim days As Integer   '����s�ڤѼ�
  deposit = InputBox("�п�J�Ĥ@�Ѧs�ڪ��B")
  days = InputBox("�п�J�s�ڤѼ�")
  Print "�Ѽ�", "��Ѧs���B", "�ֿn�s���B"  '�ѦL�����X
  For i = 1 To days    '�j�骺���Ʊq�Ĥ@�Ѩ�̫�@��
    day_deposit = deposit * 2 ^ (i - 1) '�C�Ѧs�ڬ��Ĥ@�Ѧs�ڪ�2^(i-1)��
    total_deposit = total_deposit + day_deposit  '�ֿn�`�s�ڪ��B
     Print i, Format(day_deposit, "#########,###") _
                     , Format(total_deposit, "#########,###") '�ѦL�����X
  Next i
End Sub

