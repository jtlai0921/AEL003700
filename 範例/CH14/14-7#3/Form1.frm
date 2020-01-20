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
   StartUpPosition =   3  '系統預設值
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   Const pi = 3.14159
   Circle (400, 500), 300, , 0, pi / 2               '上排第1個
   Circle (1200, 500), 300, , 0, pi                  '上排第2個
   Circle (2000, 500), 300, , 0, pi * 3 / 2          '上排第3個
   Circle (2800, 500), 300, , -pi * 2, -pi * 3 / 2   '上排第4個
   Circle (3600, 500), 300, , -pi * 2, -pi / 2       '上排第5個
   Circle (400, 1500), 300, , 0, -pi / 2                 '下排第1個
   Circle (1200, 1500), 300, , -pi * 2, pi               '下排第2個
   Circle (2000, 1500), 300, , -pi / 4, -pi * 7 / 4      '下排第3個
   Circle (2800, 1500), 300, , -pi, -pi * 3 / 2          '下排第4個
   Circle (3600, 1500), 300, , -pi * 5 / 4, -pi * 3 / 4  '下排第5個
End Sub

