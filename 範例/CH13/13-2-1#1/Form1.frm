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
   StartUpPosition =   3  '�t�ιw�]��
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   s = "I saw a saw saw a saw"
   Print InStr(s, "saw")    '�q��1�Ӧr���}�l��
   Print InStr(7, s, "saw")  '�q��7�Ӧr���}�l��
   Print InStr(s, "see")    '�䤣��
   Print InStr(30, s, "saw") '�}�l��m�w�W�L�r�����
   Print InStr(s, "")       '�n��M���r�ꬰ�Ŧr��
End Sub

