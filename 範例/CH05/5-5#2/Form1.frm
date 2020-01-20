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
   Dim name As String * 12  '宣告name為固定長度12個字元
                            '的字串變數，內容均為空白
   name = "Tom"       '將Tom接上9個空白後，再存進name
                      '中，佔12個位元組
   Print name         '顯示Tom及9個空白
   name = "Tom Jones" '將Tom Jones接上3個空白字元後，
                      '再存進name中，佔12個位元組
   Print name         '顯示Tom Jones及3個空白
   name = "張惠妹"    '將「張惠妹」接上9個空白字元後，
                      '再存進name中，佔15個位元組
   name = "我喜歡Visual BASIC" '將12個字元「我喜歡Visual BA」
                               '存進name中，佔15個位元組
   Print name         '顯示「我喜歡Visual BA」

End Sub

