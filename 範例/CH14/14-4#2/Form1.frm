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
   BackColor = RGB(255, 255, 255) '�I�����զ�
   Line (500, 200)-(1000, 700)
   Line -(100, 700)
   Line -(500, 200)
   DrawStyle = 1                  '�}��u
   Line (2000, 200)-(2500, 700)
   Line -(1500, 700)
   Line -(2000, 200)
   DrawStyle = 2                  '�I�u
   Line (3500, 200)-(4000, 700)
   Line -(3000, 700)
   Line -(3500, 200)
End Sub

