VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '�t�ιw�]��
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Dim free As Variant  '�ŧifree�����w���ܼ�
Dim dat As Date      '�ŧidat������ܼ�
free = "���w���ܼ�"  '�N�r��u�ۥ��ܼơv�s�ifree���A'��ƫ��A��String
Print TypeName(free) '���free�ثe����ƫ��A��String
Print free           '��ܦr��u�ۥ��ܼơv
free = 1234          '�N���1234�s�ifree���A��ƫ��A��Integer
Print TypeName(free) '���free�ثe����ƫ��A��Integer
Print free           '��ܾ��1234
free = 1234.56       '�N����1234.56�s�ifree���A��ƫ��A��Single
Print TypeName(free) '���free�ثe����ƫ��A��Double
Print free           '��ܳ���1234.56
free = True          '�N�޿��True�s�ifree���A��ƫ��A��Boolean
Print TypeName(free) '���free�ثe����ƫ��A��Boolean
Print free           '����޿��True
dat = "2005/12/5"    '�N����u2005/12/5�v�s�i����ܼ�dat��
free = dat           '�Ndat�ܼƪ����e�s�ifree���A��ƫ��A��Date
Print TypeName(free) '���free�ثe����ƫ��A��Date
Print free           '��ܤ���u2005/12/5�v
End Sub

