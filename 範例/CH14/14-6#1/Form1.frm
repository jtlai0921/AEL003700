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
   CurrentX = 400: CurrentY = 400  '�]�w��X�_�I���y��
   Print "(300,500)"               '��X�_�I���y��
   DrawWidth = 2                   '�]�w�u���e��
   Line (300, 500)-(1000, 1200)    '�e���u
   Print "("; CurrentX; ","; CurrentY; ")" '��X���I���y��
End Sub

