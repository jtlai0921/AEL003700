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
   StartUpPosition =   3  '�t�ιw�]��
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   a = 10: b = 20
   Print "�洫�e:"; a, b
   Call exchange(a, b)
   Print "�洫��:"; a, b
   a = "TOM": b = "JOHN"
   Print "�洫�e:"; a, b
   Call exchange(a, b)
   Print "�洫��:"; a, b
End Sub
Private Sub exchange(x, y)
  t = x
  x = y
  y = t
End Sub
