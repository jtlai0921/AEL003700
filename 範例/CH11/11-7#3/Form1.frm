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
   ReDim a(5) As Integer
   ReDim b(5) As Integer
   For i = 0 To 5
      a(i) = i * i  '�̧Ǭ�0,1,4,9,16,25
      b(i) = 5 * i  '�̧Ǭ�0,5,10,15,20,25
   Next i
   ReDim a(8)          '���s�ŧi,�M���줺�e
   ReDim Preserve b(8) '���s�ŧi,�O�s�줺�e
   For i = 0 To 8    '�C�L�}�Ca�U���������e
      Print a(i);
   Next i
   Print
   For i = 0 To 8    '�C�L�}�Cb�U���������e
      Print b(i);
   Next i
End Sub

