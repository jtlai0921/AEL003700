VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   3750
   StartUpPosition =   3  '�t�ιw�]��
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   FontSize = 16
   BackColor = QBColor(14) '����
   Print "�t�ιw�]���e�����¦�"
   CurrentX = 500: CurrentY = 500
   ForeColor = QBColor(9)  '�_�Ŧ�
   FontSize = 14
   Print "�I������,�e���_�Ŧ�"
   CurrentX = 1000: CurrentY = 1000
   ForeColor = QBColor(4)  '����
   FontSize = 12
   Print "�I������,�e������"
End Sub
