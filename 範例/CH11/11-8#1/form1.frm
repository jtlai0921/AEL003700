VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5265
   StartUpPosition =   3  '�t�ιw�]��
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   Dim tree As Variant
   Dim n As Variant
   x = Array("�Q", "��", "��")
   For Each tree In x
      Print tree '��X�}�Cx���C�@�������e
   Next
   Dim y(1, 2) As Integer
   y(0, 0) = 0: y(0, 1) = 1: y(0, 2) = 2
   y(1, 0) = 3: y(1, 1) = 4: y(1, 2) = 5
   s = 0
   For Each n In y
      Print n;    '��X�C�@���������e
      s = s + n   '�֥[�}�Cy�C�@�������M
   Next
   Print
   Print "�}�C�U�������`�M="; s
End Sub

