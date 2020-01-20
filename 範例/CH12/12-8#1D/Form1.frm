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
   Dim i As Integer
   Dim s As Double
   s = 0
   For i = 2 To 10 Step 2
      s = s + fac(i)
   Next i
   Print "SUM="; s
End Sub
Private Function fac(x As Integer) As Double
   Dim i As Integer
   fac = 1
   For i = 1 To x
      fac = fac * i
   Next i
End Function

