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
   Print , "*** �P��M�� ***"
   Print
   Print "�Ϯѽs��", "��  ��", "�ƶq", "���B"
   For i = 1 To 5
      bookno = Choose(i, 1001, 1005, 1200, 2008, 3100)
      price = Choose(i, 300, 200, 150, 100, 120)
      qty = Choose(i, 5, 10, 8, 20, 5)
      amount = price * qty              '�p��浧�Ѵ�
      totamount = totamount + amount    '�֭p�`�Ѵ�
      Print bookno, price, qty, amount  '�C�L�浧���
   Next i
   tax = totamount * 0.05               '�p����~�|
   Print
   Print "�ѴڦX�p", , , totamount      '�C�L�`�Ѵ�
   Print "��~�|(5%)", , , tax          '�C�L��~�|
   Print
   Print "*�����`�B*", , , totamount + tax  '�C�L�t�|�`�B
End Sub
