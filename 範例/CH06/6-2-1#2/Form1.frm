VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5880
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   Print "第一區", "第二區", "第三區"
   Print "1234567890123456", "第二項資料"
   Print "第一項資料", , "第二項資料"
   Print "A", "B",
   Print "C"
End Sub

