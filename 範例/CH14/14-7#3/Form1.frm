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
   Const pi = 3.14159
   Circle (400, 500), 300, , 0, pi / 2               '�W�Ʋ�1��
   Circle (1200, 500), 300, , 0, pi                  '�W�Ʋ�2��
   Circle (2000, 500), 300, , 0, pi * 3 / 2          '�W�Ʋ�3��
   Circle (2800, 500), 300, , -pi * 2, -pi * 3 / 2   '�W�Ʋ�4��
   Circle (3600, 500), 300, , -pi * 2, -pi / 2       '�W�Ʋ�5��
   Circle (400, 1500), 300, , 0, -pi / 2                 '�U�Ʋ�1��
   Circle (1200, 1500), 300, , -pi * 2, pi               '�U�Ʋ�2��
   Circle (2000, 1500), 300, , -pi / 4, -pi * 7 / 4      '�U�Ʋ�3��
   Circle (2800, 1500), 300, , -pi, -pi * 3 / 2          '�U�Ʋ�4��
   Circle (3600, 1500), 300, , -pi * 5 / 4, -pi * 3 / 4  '�U�Ʋ�5��
End Sub

