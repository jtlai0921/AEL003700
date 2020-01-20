VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1920
   BeginProperty Font 
      Name            =   "新細明體"
      Size            =   9.75
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   1185
   ScaleWidth      =   1920
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   For k = 1 To 5
      Print k;
   Next k
   Print
   Print "離開迴圈後的k值="; k
End Sub
