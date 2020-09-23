VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Login"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4215
   StartUpPosition =   2  'CenterScreen
   Begin SMIAS.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   0
      TabIndex        =   7
      Top             =   750
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdLog 
      Caption         =   "Log-in"
      Default         =   -1  'True
      Height          =   315
      Left            =   1725
      TabIndex        =   2
      Top             =   1950
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2940
      TabIndex        =   3
      Top             =   1950
      Width           =   1110
   End
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -600
      TabIndex        =   6
      Top             =   1800
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   53
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   975
      MaxLength       =   100
      PasswordChar    =   "*"
      TabIndex        =   1
      Tag             =   "Name"
      Top             =   1350
      Width           =   1515
   End
   Begin MSDataListLib.DataCombo dcUser 
      Height          =   315
      Left            =   975
      TabIndex        =   0
      Top             =   975
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please select your username and enter your password in the space provided bellow."
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   750
      TabIndex        =   8
      Top             =   150
      Width           =   3315
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   765
      Left            =   675
      Top             =   0
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   780
      Left            =   0
      Picture         =   "frmLogin.frx":038A
      Top             =   0
      Width           =   720
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Password:"
      Height          =   240
      Index           =   1
      Left            =   75
      TabIndex        =   5
      Top             =   1350
      Width           =   840
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Username:"
      Height          =   240
      Index           =   18
      Left            =   -300
      TabIndex        =   4
      Top             =   975
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub cmdCancel_Click()
    MAIN.CloseMe = True
    Unload Me
End Sub

Private Sub cmdLog_Click()
    'Verify
    If dcUser.Text = "" Then dcUser.SetFocus: Exit Sub
    If txtPass.Text = "" Then txtPass.SetFocus: Exit Sub
    Dim strPass As String
    strPass = getValueAt("SELECT PK,Password FROM tbl_SM_Users WHERE PK=" & dcUser.BoundText, "Password")
    strPass = Enc.DecryptString(strPass)
    'Very short code of login system
    If LCase(txtPass.Text) = LCase(strPass) Then
        With CurrUser
            .USER_NAME = dcUser.Text
            .USER_PK = dcUser.BoundText
            .USER_ISADMIN = CBool(changeYNValue(getValueAt("SELECT PK,Admin FROM tbl_SM_Users WHERE PK=" & dcUser.BoundText, "Admin")))
        End With
        Unload Me
    Else
        MsgBox "Invalid password.Please try again!", vbExclamation
        txtPass.SetFocus
    End If
    strPass = vbNullString

End Sub

Private Sub Form_Load()
    bind_dc "SELECT * FROM tbl_SM_Users", "UserID", dcUser, "PK"
End Sub

Private Sub txtPass_Change()
    txtPass.SelStart = Len(txtPass.Text)
End Sub

Private Sub txtPass_GotFocus()
    HLText txtPass
End Sub
