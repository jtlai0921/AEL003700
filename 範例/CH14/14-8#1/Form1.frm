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
  FillStyle = 0                       '���
  Line (400, 400)-(1000, 1000), , B   '�e�W�ƲĤ@�Ӥ��
  FillStyle = 1                       '�z��
  Line (1400, 400)-(2000, 1000), , B  '�e�W�ƲĤG�Ӥ��
  FillStyle = 2                       '�����u
  Line (2400, 400)-(3000, 1000), , B  '�e�W�ƲĤT�Ӥ��
  FillStyle = 3                       '�����u
  Line (3400, 400)-(4000, 1000), , B  '�e�W�Ʋĥ|�Ӥ��
  FillStyle = 4                       '���W��k�U���׽u
  Line (400, 1600)-(1000, 2200), , B  '�e�U�ƲĤ@�Ӥ��
  FillStyle = 5                       '���U��k�W���׽u
  Line (1400, 1600)-(2000, 2200), , B  '�e�U�ƲĤG�Ӥ��
  FillStyle = 6                        '������e�u
  Line (2400, 1600)-(3000, 2200), , B  '�e�U�ƲĤT�Ӥ��
  FillStyle = 7                        '�﨤��e�u
  Line (3400, 1600)-(4000, 2200), , B  '�e�U�Ʋĥ|�Ӥ��
End Sub

