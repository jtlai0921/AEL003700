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
  Dim code As Byte      '�ŧicode���줸���ܼơA�|����1�Ӧ줸�ժ��x�s��m
  Dim year As Integer   '�ŧiyear������ܼơA�|����2 �Ӧ줸�ժ��x�s��m
  Dim salary As Long    '�ŧisalary��������ܼơA�|����4 �Ӧ줸�ժ��x�s��m
  Dim price As Single   '�ŧiprice�������ܼơA�|����4 �Ӧ줸�ժ��x�s��m
  Dim qty As Double     '�ŧiqty��������ܼơA�|����8 �Ӧ줸�ժ��x�s��m
  Dim name As String    '�ŧiname���ܰʪ��ת��r���ܼơA�����x�s��m�|�H�ۤ��e����
  Dim address As String * 30 '�ŧiaddress���T�w���ת��r���ܼơA�T�w����30�Ӧ줸��
End Sub

