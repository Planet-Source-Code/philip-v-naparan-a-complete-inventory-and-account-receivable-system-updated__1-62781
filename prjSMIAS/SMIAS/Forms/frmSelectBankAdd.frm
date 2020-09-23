VERSION 5.00
Begin VB.Form frmSelectBankAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Information"
   ClientHeight    =   2850
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
   Icon            =   "frmSelectBankAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowser 
      Height          =   315
      Left            =   2755
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Browse..."
      Top             =   1600
      Width           =   315
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1230
      MaxLength       =   150
      TabIndex        =   2
      Top             =   900
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1230
      MaxLength       =   150
      TabIndex        =   3
      Top             =   1260
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1230
      MaxLength       =   10
      TabIndex        =   4
      Top             =   1620
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2790
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   1590
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
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
      Left            =   -225
      TabIndex        =   13
      Top             =   2025
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   53
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "City/Town"
      Height          =   255
      Left            =   75
      TabIndex        =   12
      Top             =   900
      Width           =   1050
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      Caption         =   "Zip Code"
      Height          =   255
      Left            =   75
      TabIndex        =   11
      Top             =   1650
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Province"
      Height          =   255
      Left            =   75
      TabIndex        =   10
      Top             =   1275
      Width           =   1020
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Address"
      Height          =   255
      Left            =   75
      TabIndex        =   9
      Top             =   525
      Width           =   1020
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Bank Name"
      Height          =   255
      Left            =   75
      TabIndex        =   8
      Top             =   150
      Width           =   1050
   End
End
Attribute VB_Name = "frmSelectBankAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ADD_STATE        As Boolean
Public PK         As Long

Private Sub cmdBrowser_Click()
    With frmSelectZipCode
        Set .txtCity = Text3
        Set .txtState = Text4
        Set .txtZipCode = Text5
    
        .show vbModal
    End With
End Sub

Private Sub Command1_Click()
    If is_empty(Text1) = True Then Exit Sub
    If is_empty(Text2) = True Then Exit Sub

    If ADD_STATE = False Then
        If isRecordExist("tbl_SM_BankList", "PK", PK) = False Then
            MsgBox "This bank is no longer exist in the record. Click ok to reload the records!", vbExclamation, "Unable To Edit"
            frmSelectBank.reload_rec
            Unload Me
            Exit Sub
        End If
    End If
    
    On Error GoTo err
    With frmSelectBank.rs
        If ADD_STATE = True Then: .AddNew: .Fields("PK") = PK
            .Fields("BankName") = Text1.Text
            .Fields("Address") = Text2.Text
            .Fields("CityTown") = Text3.Text
            .Fields("Province") = Text4.Text
            .Fields("ZipCode") = Text5.Text
        .Update
    End With
    frmSelectBank.reload_rec
    Unload Me
    Exit Sub
err:
        prompt_err err, Me.Name, "Command1_Click"
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Set the graphics for the controls
    With MAIN
        cmdBrowser.Picture = .i16x16.ListImages(8).Picture
    End With
    
    If ADD_STATE = True Then
       Caption = "Add New"
       PK = getIndex("tbl_SM_BankList")
    Else
        Caption = "Edit Existing"
        
        customMove frmSelectBank.rs, False, PK, "PK"
        With frmSelectBank.rs
            Text1.Text = .Fields("BankName")
            Text2.Text = .Fields("Address")
            Text3.Text = .Fields("CityTown")
            Text4.Text = .Fields("Province")
            Text5.Text = .Fields("ZipCode")
        End With
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSelectBankAdd = Nothing
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

Private Sub Text4_GotFocus()
    HLText Text4
End Sub

Private Sub Text5_GotFocus()
    HLText Text5
End Sub
