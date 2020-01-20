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
   Dim name As String   '宣告name為變動長度，開始時佔用0個字元
   Print Len(name)      '輸出字串變數name的長度（字元數）
   name = "Tom"       '將Tom存進name，3個字元佔用3個位元組
   Print Len(name)      '輸出字串變數name的長度（字元數）
   name = "Tom Jones"  '將Tom Jones存進name，9個字元佔用9個位元組
   Print Len(name)      '輸出字串變數name的長度（字元數）
   name = "張惠妹"    '將「張惠妹」存進name，3個字元佔用6個位元組
   Print Len(name)      '輸出字串變數name的長度（字元數）
   name = "我喜歡Visual BASIC" '將「我喜歡Visual BASIC」存進name中，15個字元佔18個位元組
   Print Len(name)      '輸出字串變數name的長度（字元數），結果為15
End Sub

