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
   CurrentX = 1000
   CurrentY = 500
   Print "���q��r���W���y�Ь�1000,500"
   CurrentX = 2000
   CurrentY = 1000
   Print "���q��r���W���y�Ь�2000,1000"
   CurrentX = 500
   CurrentY = 2000
   Print "���q��r���W���y�Ь�500,2000"
End Sub

