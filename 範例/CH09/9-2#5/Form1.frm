VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   3750
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   Print "*****  銷售數量統計圖  *****"
   Print
   Print "星期      銷  售  數  量"
   Print
   For i = 1 To 6
      qty = Choose(i, 10, 7, 8, 12, 15, 20)
      Print i; Tab(8);
      For j = 1 To qty
         Print "*";
      Next j
      Print
   Next i
End Sub

