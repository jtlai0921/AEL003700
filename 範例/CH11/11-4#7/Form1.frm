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
Option Base 1

Private Sub Form_Activate()
   Dim a(3, 4)
   For i = 1 To 3
      For j = 1 To 4
         a(i, j) = InputBox("�п�J�@�Ӽ�")
      Next j
   Next i
   Rem �p��ΦC�L�`�M
   s = 0
   For i = 1 To 4
      Print a(2, i),
      s = s + a(2, i)
   Next i
   Print
   Print "��2�C�������`�M="; s
End Sub

