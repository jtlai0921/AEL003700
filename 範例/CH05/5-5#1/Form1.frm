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
   Dim name As String   '�ŧiname���ܰʪ��סA�}�l�ɦ���0�Ӧr��
   Print Len(name)      '��X�r���ܼ�name�����ס]�r���ơ^
   name = "Tom"       '�NTom�s�iname�A3�Ӧr������3�Ӧ줸��
   Print Len(name)      '��X�r���ܼ�name�����ס]�r���ơ^
   name = "Tom Jones"  '�NTom Jones�s�iname�A9�Ӧr������9�Ӧ줸��
   Print Len(name)      '��X�r���ܼ�name�����ס]�r���ơ^
   name = "�i�f�f"    '�N�u�i�f�f�v�s�iname�A3�Ӧr������6�Ӧ줸��
   Print Len(name)      '��X�r���ܼ�name�����ס]�r���ơ^
   name = "�ڳ��wVisual BASIC" '�N�u�ڳ��wVisual BASIC�v�s�iname���A15�Ӧr����18�Ӧ줸��
   Print Len(name)      '��X�r���ܼ�name�����ס]�r���ơ^�A���G��15
End Sub

