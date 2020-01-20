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
Option Base 1
Private Sub Form_Activate()
   Dim s()
   s = Array(80, 90, 70, 82, 76)
   Print "平均值為"; average(s())
End Sub
Private Function average(array1())
   total = 0
   n = UBound(array1)
   For i = 1 To n
      total = total + array1(i)
   Next i
   average = total / n
End Function

