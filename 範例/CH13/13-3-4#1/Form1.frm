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
   a = "123"  '�r����
   b = 123    '�ƭȸ��
   Print a    '��X�r��
   Print b    '��X�ƭ�
   Print Str(b)      '�ƭ���r��
   Print Len(Str(b)) '��X�r�����
   Print "3.141593�����׬O:"; Len(Str(3.141593)) - 1
End Sub

