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
  Dim code As Byte      '宣告code為位元組變數，會佔用1個位元組的儲存位置
  Dim year As Integer   '宣告year為整數變數，會佔用2 個位元組的儲存位置
  Dim salary As Long    '宣告salary為長整數變數，會佔用4 個位元組的儲存位置
  Dim price As Single   '宣告price為單精數變數，會佔用4 個位元組的儲存位置
  Dim qty As Double     '宣告qty為倍精數變數，會佔用8 個位元組的儲存位置
  Dim name As String    '宣告name為變動長度的字串變數，佔用儲存位置會隨著內容改變
  Dim address As String * 30 '宣告address為固定長度的字串變數，固定佔用30個位元組
End Sub

