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
   Print "�{���}�l"
   Call subprog1
   Print "�{������"
End Sub
Private Sub subprog1()
   Print " *�Ƶ{��1�}�l"
   Call subprog2
   Print " *�Ƶ{��1����"
End Sub
Private Sub subprog2()
   Print "  **�Ƶ{��2�}�l"
   Print "  **�Ƶ{��2����"
End Sub
