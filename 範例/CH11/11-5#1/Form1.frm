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
   Dim a(5), b(3, 4), c(-3 To 3, 5, 2 To 6)
   Print LBound(a); UBound(a)
   Print
   Print LBound(b, 1); UBound(b, 1)
   Print LBound(b, 2); UBound(b, 2)
   Print
   For i = 1 To 3
      Print LBound(c, i); UBound(c, i)
   Next i
End Sub

