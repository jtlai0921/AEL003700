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
  FillStyle = 0                       '實心
  Line (400, 400)-(1000, 1000), , B   '畫上排第一個方框
  FillStyle = 1                       '透明
  Line (1400, 400)-(2000, 1000), , B  '畫上排第二個方框
  FillStyle = 2                       '水平線
  Line (2400, 400)-(3000, 1000), , B  '畫上排第三個方框
  FillStyle = 3                       '垂直線
  Line (3400, 400)-(4000, 1000), , B  '畫上排第四個方框
  FillStyle = 4                       '左上到右下的斜線
  Line (400, 1600)-(1000, 2200), , B  '畫下排第一個方框
  FillStyle = 5                       '左下到右上的斜線
  Line (1400, 1600)-(2000, 2200), , B  '畫下排第二個方框
  FillStyle = 6                        '垂直交叉線
  Line (2400, 1600)-(3000, 2200), , B  '畫下排第三個方框
  FillStyle = 7                        '對角交叉線
  Line (3400, 1600)-(4000, 2200), , B  '畫下排第四個方框
End Sub

