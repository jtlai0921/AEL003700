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
   a = "  我愛Visual BASIC  "  '前後各有二個空格
   b = LTrim(a)
   c = RTrim(a)
   d = Trim(a)
   Print a, Len(a)
   Print b, Len(b)
   Print c, Len(c)
   Print d, Len(d)
End Sub
