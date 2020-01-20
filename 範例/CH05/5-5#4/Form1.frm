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
   Dim date1 As Date, date2 As Date '宣告date1與date2均為日期變數
   date1 = #12/25/2005#    '將2005年12月25日存進Date1
   date2 = #1/5/2006#      '將2006年1月5日存進Date2
   Print date1              '顯示2005/12/25
   Print date2              '顯示2006/1/5
   Print date2 - date1       '顯示11（兩個日期相差的日數）
   Print date1 + 11         '顯示2006/1/5
   Print date2 - 11         '顯示2005/12/25
End Sub

