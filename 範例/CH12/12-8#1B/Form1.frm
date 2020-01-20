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
   Dim i As Integer, j As Integer
   Dim p As Double, s As Double
   s = 0
   For i = 2 To 10 Step 2
      GoSub fac
      s = s + p
   Next i
   Print "SUM="; s
   Exit Sub
fac:
   p = 1
   For j = 1 To i
      p = p * j
   Next j
   Return
End Sub

