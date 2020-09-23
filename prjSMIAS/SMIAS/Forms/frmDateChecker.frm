VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDateChecker 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Current Date and Time Checker"
   ClientHeight    =   4110
   ClientLeft      =   255
   ClientTop       =   1695
   ClientWidth     =   3480
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
   Icon            =   "frmDateChecker.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOK 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   390
      Left            =   1125
      TabIndex        =   5
      Top             =   3525
      Width           =   1290
   End
   Begin VB.Timer tmrCurrDate 
      Interval        =   1000
      Left            =   2700
      Top             =   2100
   End
   Begin VB.Timer tmrCurrTime 
      Interval        =   1000
      Left            =   3300
      Top             =   2325
   End
   Begin VB.CheckBox Check1 
      Caption         =   "No,Let me adjust it!"
      Height          =   315
      Left            =   225
      TabIndex        =   0
      Top             =   1125
      Width           =   1740
   End
   Begin VB.CheckBox Check2 
      Caption         =   "No,Let me adjust it!"
      Height          =   315
      Left            =   225
      TabIndex        =   2
      Top             =   2325
      Width           =   1740
   End
   Begin VB.Timer tmrBlink 
      Interval        =   500
      Left            =   2925
      Top             =   825
   End
   Begin VB.Frame Frames 
      Enabled         =   0   'False
      Height          =   915
      Index           =   1
      Left            =   75
      TabIndex        =   3
      Top             =   2400
      Width           =   3315
      Begin MSComCtl2.DTPicker dpTime 
         Height          =   390
         Left            =   150
         TabIndex        =   4
         Top             =   375
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   24576002
         CurrentDate     =   38557
      End
   End
   Begin VB.Frame Frames 
      Enabled         =   0   'False
      Height          =   915
      Index           =   0
      Left            =   75
      TabIndex        =   1
      Top             =   1125
      Width           =   3315
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   390
         Left            =   150
         TabIndex        =   9
         Top             =   375
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   0
         CalendarTitleBackColor=   8388608
         CalendarTitleForeColor=   16777215
         CalendarTrailingForeColor=   12632256
         CustomFormat    =   "MMMM dd, yyyy"
         Format          =   24576003
         CurrentDate     =   38207
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   75
      Picture         =   "frmDateChecker.frx":000C
      Top             =   150
      Width           =   480
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Please make sure that the current time and date settings are correct!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   765
      Index           =   0
      Left            =   675
      TabIndex        =   8
      Top             =   150
      Width           =   2640
   End
   Begin VB.Label Labels 
      Caption         =   "Is this correct?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   75
      TabIndex        =   6
      Top             =   900
      Width           =   1665
   End
   Begin VB.Label Labels 
      Caption         =   "Is this correct?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   75
      TabIndex        =   7
      Top             =   2100
      Width           =   1665
   End
End
Attribute VB_Name = "frmDateChecker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub btnOK_Click()
    If Check1.Value = 1 Then Date = dtpDate.Value
    If Check2.Value = 1 Then Time = dpTime.Value
    Unload Me
End Sub

Private Sub Check1_Click()
    DisplayCap
    If Check1.Value = 1 Then
        Frames(0).Enabled = True
        tmrCurrDate.Enabled = False
    Else
        Frames(0).Enabled = False
        dtpDate.Value = Date
        tmrCurrDate.Enabled = True
    End If
End Sub

Private Sub Check2_Click()
    DisplayCap
    If Check2.Value = 1 Then
        Frames(1).Enabled = True
        tmrCurrTime.Enabled = False
    Else
        Frames(1).Enabled = False
        tmrCurrTime.Enabled = True
        dpTime.Value = Time
    End If
End Sub

Private Sub Form_Load()
    dtpDate.Value = Date
    dpTime.Value = Time
End Sub

Private Sub tmrBlink_Timer()
    Labels(0).Visible = Not Labels(0).Visible
End Sub

Private Sub tmrCurrDate_Timer()
    If dtpDate.Value <> Date Then dtpDate.Value = Date
End Sub

Private Sub tmrCurrTime_Timer()
    dpTime.Value = Time
End Sub

Private Sub DisplayCap()
    If Check1.Value = 1 Or Check2.Value = 1 Then
        btnOK.Caption = "Adjust"
    Else
        btnOK.Caption = "Close"
    End If
End Sub
