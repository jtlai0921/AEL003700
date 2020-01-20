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
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   Print "程式開始"
   Call subprog1
   Print "程式結束"
End Sub
Private Sub subprog1()
   Print " *副程式1開始"
   Call subprog2
   Print " *副程式1結束"
End Sub
Private Sub subprog2()
   Print "  **副程式2開始"
   Print "  **副程式2結束"
End Sub
