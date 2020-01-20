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
   StartUpPosition =   3  '╰参箇砞
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   Dim a(4) As Integer
   For i = 0 To 4
      a(i) = Val(InputBox("叫块材" + Str(i + 1) + "计"))
   Next i
   s = a(0) + a(1) + a(2) + a(3) + a(4)
   Print "场计羆㎝="; s
End Sub

