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
   BackColor = &HFFFFFF         '白色
   ForeColor = QBColor(9)       '藍色
   FillColor = RGB(255, 255, 0) '黃色
   FillStyle = 0
   Print "電腦繪圖也可輸出文字"
   ForeColor = &HFF             '紅色
   Line (500, 800)-(1500, 1500), , B
   ForeColor = RGB(0, 255, 0)   '綠色
   FillStyle = 1
   Line (2000, 800)-(3000, 1500), , B
End Sub

