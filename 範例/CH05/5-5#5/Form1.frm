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
Dim free As Variant  '宣告free為不定型變數
Dim dat As Date      '宣告dat為日期變數
free = "不定型變數"  '將字串「自由變數」存進free中，'其副型態為String
Print TypeName(free) '顯示free目前的資料型態為String
Print free           '顯示字串「自由變數」
free = 1234          '將整數1234存進free中，其副型態為Integer
Print TypeName(free) '顯示free目前的資料型態為Integer
Print free           '顯示整數1234
free = 1234.56       '將單精數1234.56存進free中，其副型態為Single
Print TypeName(free) '顯示free目前的資料型態為Double
Print free           '顯示單精數1234.56
free = True          '將邏輯值True存進free中，其副型態為Boolean
Print TypeName(free) '顯示free目前的資料型態為Boolean
Print free           '顯示邏輯值True
dat = "2005/12/5"    '將日期「2005/12/5」存進日期變數dat中
free = dat           '將dat變數的內容存進free中，其副型態為Date
Print TypeName(free) '顯示free目前的資料型態為Date
Print free           '顯示日期「2005/12/5」
End Sub

