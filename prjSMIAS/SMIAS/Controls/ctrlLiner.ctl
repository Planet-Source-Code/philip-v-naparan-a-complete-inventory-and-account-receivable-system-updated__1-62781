VERSION 5.00
Begin VB.UserControl ctrlLiner 
   ClientHeight    =   1425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7155
   ScaleHeight     =   1425
   ScaleWidth      =   7155
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      X1              =   675
      X2              =   5025
      Y1              =   225
      Y2              =   225
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   4425
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "ctrlLiner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''*****************************************************************
'' File Name: ctrlLiner.ctl
'' Purpose: Control used to draw a border line
'' Required Files: NONE
''
'' Programmer: Philip V. Naparan   E-mail: philipnaparan@yahoo.com
'' Date Created: Dec-20-04 1:55 AM
'' Last Modified:
'' Modified By:
'' Credits: NONE, ALL CODES ARE CODED BY Philip V. Naparan
''*****************************************************************

Option Explicit

Private Sub UserControl_Initialize()
    On Error Resume Next
    UserControl.Height = 30
    UserControl.BackColor = UserControl.Parent.BackColor
End Sub

Private Sub UserControl_InitProperties()
        UserControl.Height = 30
    MsgBox "Developed By: Philip V. Naparan" & vbCrLf & _
       "Website: http://www.naparansoft.com or http://www.Your Company.com" & vbCrLf & _
       "E-mail: philipnaparan@yahoo.com" & vbCrLf & vbCrLf & _
       "100% GAWANG PINOY", vbInformation, "About"
End Sub

Private Sub UserControl_Paint()
'*** bellow code can be use also
'    UserControl.Height = 30
'    UserControl.Line (0, 0)-(UserControl.Width, 0), &H80000010
'    UserControl.Line (0, 20)-(UserControl.Width, 20), &H80000014
Line1.X1 = 0
Line1.Y1 = 0
Line1.X2 = UserControl.Width
Line1.Y2 = 0

Line2.X1 = 0
Line2.Y1 = 20
Line2.X2 = UserControl.Width
Line2.Y2 = 20
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 30
    UserControl_Paint
End Sub
