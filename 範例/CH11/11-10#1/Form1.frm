VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   4680
   StartUpPosition =   3  '�t�ιw�]��
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1  '���ԭz�n�w�Ʀb(�@��)(�ŧi)��
Private Sub Form_Activate()
   Rem �]�w�ʺA�}�C�ο�J�䤺�e
   Dim n() As String  '�ŧi���s��m�W���ʺA�}�C
   Dim s() As Single  '�ŧi���s����ƪ��ʺA�}�C
   rec = 0            '�ΨӰO����J������
   nam = InputBox$("�п�J��1�Ӿǥͪ��m�W")
   Do While nam <> "end"    '��J�m�W���Oend�N�~��j��
      rec = rec + 1         '��J���ƥ[1
      score = Val(InputBox("�п�J��" + Str(rec) + "�Ӿǥͪ�����"))
      ReDim Preserve n(rec)
      ReDim Preserve s(rec)
      n(rec) = nam
      s(rec) = score
      nam = InputBox$("�п�J��" + Str(rec + 1) + "�Ӿǥͪ��m�W")
   Loop
   Rem ��X���G
   Print "    ��  ��  ��  �Z  ��"
   Print "�m  �W", "��  �Z"
   For i = 1 To rec
      Print n(i), s(i)
   Next i
End Sub

