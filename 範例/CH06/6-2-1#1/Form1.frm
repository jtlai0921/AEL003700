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
   Print "第一區", "第二區", "第三區"
   Print 1, 2, 3
   Print -4, -5
   FontSize = 12
   Print "第一區", "第二區", "第三區"
   Print 1, 2
   Print -3, -4, -5
   Width = 6000
   FontSize = 14
   Print "第一區", "第二區", "第三區"
   Print "Visual", "BASIC", "真的很棒"
End Sub

