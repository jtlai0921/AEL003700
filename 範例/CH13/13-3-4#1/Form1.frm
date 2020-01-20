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
   a = "123"  '字串資料
   b = 123    '數值資料
   Print a    '輸出字串
   Print b    '輸出數值
   Print Str(b)      '數值轉字串
   Print Len(Str(b)) '輸出字串長度
   Print "3.141593的長度是:"; Len(Str(3.141593)) - 1
End Sub

