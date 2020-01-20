VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5265
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   x = Array("松", "竹", "梅")
   Print x(0), x(1), x(2)
   y = Array(10, 20, 30, 40)
   Print y(0), y(1), y(2), y(3)
   Print y(0) + y(1) + y(2) + y(3)
End Sub

