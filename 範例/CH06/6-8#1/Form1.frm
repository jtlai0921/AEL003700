VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   3750
   StartUpPosition =   3  '�t�ιw�]��
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   FontSize = 14    '�]�w�r���j�p��14�I
   Print Now       '��X����P�ɶ�
   Print Date       '��X���
   Print Time       '��X�ɶ�
   Print            '�Ť@�C
   Print "����";
   Print Val(Format(Date, "yyyy")) - 1911; _
         "�~";
   Print Format(Date, "m"); "��";
   Print Format(Date, "d"); "��";
   Print Format(Time, "h"); "��";
   Print Format(Time, "n"); "��";
   Print Format(Time, "s"); "��";
End Sub
