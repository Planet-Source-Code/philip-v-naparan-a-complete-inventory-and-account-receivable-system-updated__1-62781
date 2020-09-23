VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4500
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   7485
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":038A
   ScaleHeight     =   4500
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrUnload 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   5325
      Top             =   1125
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

''*****************************************************************
'' File Name:
'' Purpose:
'' Required Files:
''
'' Programmer: Philip V. Naparan   E-mail: philipnaparan@yahoo.com
'' Date Created:
'' Last Modified:
'' Modified By:
'' Credits:NONE, All codes are coded by Philip V. Naparan
''*****************************************************************

Option Explicit

Public DisableLoader As Boolean
Dim c As Integer

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    If DisableLoader = False Then
        tmrUnload.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    c = 0
End Sub



Private Sub tmrUnload_Timer()
    c = c + 1
    If c = 5 Then MAIN.Enabled = True: Unload Me
    If c = 2 Then MAIN.Visible = True
End Sub
