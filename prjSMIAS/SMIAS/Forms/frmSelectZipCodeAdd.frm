VERSION 5.00
Begin VB.Form frmSelectZipCodeAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zip Code"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4050
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelectZipCodeAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2790
      TabIndex        =   4
      Top             =   1455
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   1590
      TabIndex        =   3
      Top             =   1455
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1230
      MaxLength       =   10
      TabIndex        =   2
      Top             =   870
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1230
      MaxLength       =   150
      TabIndex        =   1
      Top             =   510
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1230
      MaxLength       =   150
      TabIndex        =   0
      Top             =   150
      Width           =   2655
   End
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -150
      TabIndex        =   8
      Top             =   1275
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   53
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Province"
      Height          =   255
      Left            =   75
      TabIndex        =   7
      Top             =   525
      Width           =   1020
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      Caption         =   "Zip Code"
      Height          =   255
      Left            =   75
      TabIndex        =   6
      Top             =   900
      Width           =   1050
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "City/Town"
      Height          =   255
      Left            =   75
      TabIndex        =   5
      Top             =   150
      Width           =   1050
   End
End
Attribute VB_Name = "frmSelectZipCodeAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ADD_STATE        As Boolean
Public CURR_ZIP         As String

Private Sub Command1_Click()
    If is_empty(Text1) = True Then Exit Sub

    If ADD_STATE = True Then
        If isRecordExist("tbl_SM_ZipCodeList", "ZipCode", Text3.Text, True) = True Then
            MsgBox "Zip Code is already exist in the record. Please change it!", vbExclamation
            Text3.SetFocus
            Exit Sub
        End If
    Else
        If LCase(Text3.Text) <> LCase(CURR_ZIP) Then
            If isRecordExist("tbl_SM_ZipCodeList", "ZipCode", Text3.Text, True) = True Then
                MsgBox "Zip Code is already exist in the record. Please change it!", vbExclamation
                Text3.SetFocus
                Exit Sub
            End If
        Else
            If isRecordExist("tbl_SM_ZipCodeList", "ZipCode", CURR_ZIP, True) = False Then
                MsgBox "This zip code is no longer exist in the record. Click ok to reload the records!", vbExclamation, "Unable To Edit"
                frmSelectZipCode.reload_rec
                Unload Me
            Exit Sub
        End If
        End If
    End If
    
    On Error GoTo err
    With frmSelectZipCode.rs
        If ADD_STATE = True Then: .AddNew:
            .Fields("CityTown") = Text1.Text
            .Fields("Province") = Text2.Text
            .Fields("ZipCode") = Text3.Text
        .Update
    End With
    frmSelectZipCode.reload_rec
    Unload Me
    Exit Sub
err:
        prompt_err err, "frmSelectZipCode", "Command1_Click"
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If ADD_STATE = True Then
       Caption = "Add New"
    Else
        Caption = "Edit Existing"
        
        customMove frmSelectZipCode.rs, False, CURR_ZIP, "ZipCode"
        With frmSelectZipCode.rs
            Text1.Text = .Fields("CityTown")
            Text2.Text = .Fields("Province")
            Text3.Text = .Fields("ZipCode")
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSelectZipCodeAdd = Nothing
End Sub

Private Sub Text1_GotFocus()
    HLText Text1
End Sub

Private Sub Text2_GotFocus()
    HLText Text2
End Sub

Private Sub Text3_GotFocus()
    HLText Text3
End Sub
