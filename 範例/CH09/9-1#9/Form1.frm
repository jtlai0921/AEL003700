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
   Print , "*** 銷售清單 ***"
   Print
   Print "圖書編號", "單  價", "數量", "金額"
   For i = 1 To 5
      bookno = Choose(i, 1001, 1005, 1200, 2008, 3100)
      price = Choose(i, 300, 200, 150, 100, 120)
      qty = Choose(i, 5, 10, 8, 20, 5)
      amount = price * qty              '計算單筆書款
      totamount = totamount + amount    '累計總書款
      Print bookno, price, qty, amount  '列印單筆資料
   Next i
   tax = totamount * 0.05               '計算營業稅
   Print
   Print "書款合計", , , totamount      '列印總書款
   Print "營業稅(5%)", , , tax          '列印營業稅
   Print
   Print "*應收總額*", , , totamount + tax  '列印含稅總額
End Sub
