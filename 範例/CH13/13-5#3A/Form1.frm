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
  Do
    n = Val(InputBox("�п�J�@�ӥ����"))
    If n > 0 And n = Int(n) Then
       Exit Do
    Else
       MsgBox "�D�����,�ЦA���s��J"
    End If
  Loop
  Print "���"; n; "���]�Ʀ�:"
  For i = 1 To n
    If n / i = Int(n / i) Then Print i;
  Next i
End Sub

