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
   Dim a()
   a = Array(-12.5, -13.5, 12.5, 13.5, 12.52, 12.49)
   Print "  X", "Fix(X)", "Int(X)", "CInt(x)"
   For i = 0 To 5
     Print a(i), Fix(a(i)), Int(a(i)), CInt(a(i))
   Next i
End Sub
