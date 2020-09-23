VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Record"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "Find"
      Default         =   -1  'True
      Height          =   315
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   1251
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4320
      TabIndex        =   2
      Top             =   480
      Width           =   1251
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   5415
   End
   Begin VB.Image imgICO 
      Height          =   480
      Left            =   120
      Picture         =   "frmFind.frx":058A
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Find What:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public srcListView As ListView

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If IsEmpty(txtEntry) = True Then Exit Sub
    Call search_in_listview(srcListView, txtEntry.Text)
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFind = Nothing
End Sub

Private Sub txtEntry_GotFocus()
    HLText txtEntry
End Sub
