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
   x = 10: y = 20
   Print "呼叫副程式Sub1前:"; x; y
   Call sub1
   Print "呼叫副程式Sub1後:"; x; y
End Sub
Private Sub sub1()
   x = 4: y = 9: z = 16
   Print "    在副程式Sub1中:"; x; y
End Sub

