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
   Dim name As String * 12  '�ŧiname���T�w����12�Ӧr��
                            '���r���ܼơA���e�����ť�
   name = "Tom"       '�NTom���W9�Ӫťի�A�A�s�iname
                      '���A��12�Ӧ줸��
   Print name         '���Tom��9�Ӫť�
   name = "Tom Jones" '�NTom Jones���W3�Ӫťզr����A
                      '�A�s�iname���A��12�Ӧ줸��
   Print name         '���Tom Jones��3�Ӫť�
   name = "�i�f�f"    '�N�u�i�f�f�v���W9�Ӫťզr����A
                      '�A�s�iname���A��15�Ӧ줸��
   name = "�ڳ��wVisual BASIC" '�N12�Ӧr���u�ڳ��wVisual BA�v
                               '�s�iname���A��15�Ӧ줸��
   Print name         '��ܡu�ڳ��wVisual BA�v

End Sub

