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
   Dim date1 As Date, date2 As Date '�ŧidate1�Pdate2��������ܼ�
   date1 = #12/25/2005#    '�N2005�~12��25��s�iDate1
   date2 = #1/5/2006#      '�N2006�~1��5��s�iDate2
   Print date1              '���2005/12/25
   Print date2              '���2006/1/5
   Print date2 - date1       '���11�]��Ӥ���ۮt����ơ^
   Print date1 + 11         '���2006/1/5
   Print date2 - 11         '���2005/12/25
End Sub

