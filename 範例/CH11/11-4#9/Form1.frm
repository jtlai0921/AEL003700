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
   Dim b(4, 4)
   For i = 1 To 4
      For j = 1 To 4
         b(i, j) = InputBox("�п�J�@�Ӽ�")
      Next j
   Next i
   Rem �C�L�﨤�u�W�U���������e
   For i = 1 To 4
      Print b(i, i),
   Next i
   Print
   For i = 1 To 4
      Print b(i, 5 - i),
   Next i
End Sub

