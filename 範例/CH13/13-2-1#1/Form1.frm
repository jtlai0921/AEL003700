VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   s = "I saw a saw saw a saw"
   Print InStr(s, "saw")    '從第1個字元開始找
   Print InStr(7, s, "saw")  '從第7個字元開始找
   Print InStr(s, "see")    '找不到
   Print InStr(30, s, "saw") '開始位置已超過字串長度
   Print InStr(s, "")       '要找尋的字串為空字串
End Sub

