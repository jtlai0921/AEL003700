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
   d = InputBox("�п�J1~7���Ʀr")
   If d < 1 Or d >= 8 Then Print ("��J�Ʀr�W�X�d��") _
      Else Print Choose(d, "Monday", "Tuesday", _
             "Wednesday", "Thursday", "Friday", _
             "Saturday", "Sunday")
End Sub

 
