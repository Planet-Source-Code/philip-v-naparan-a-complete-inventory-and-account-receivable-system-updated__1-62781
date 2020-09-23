VERSION 5.00
Begin VB.Form frmUserRecAE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserRecAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   5610
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Admin"
      Height          =   240
      Left            =   1275
      TabIndex        =   3
      Top             =   1350
      Width           =   990
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1275
      MaxLength       =   100
      TabIndex        =   2
      Tag             =   "Complete Name"
      Top             =   975
      Width           =   3840
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   0
      Left            =   1275
      MaxLength       =   100
      TabIndex        =   0
      Tag             =   "Username"
      Top             =   150
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1275
      MaxLength       =   100
      TabIndex        =   1
      Tag             =   "Password"
      Top             =   525
      Width           =   2415
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   2700
      TabIndex        =   5
      Top             =   2025
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4140
      TabIndex        =   6
      Top             =   2025
      Width           =   1335
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   150
      TabIndex        =   4
      Top             =   2025
      Width           =   1680
   End
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -150
      TabIndex        =   7
      Top             =   1875
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   53
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Complete Name"
      Height          =   240
      Index           =   2
      Left            =   75
      TabIndex        =   10
      Top             =   975
      Width           =   1140
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Password"
      Height          =   240
      Index           =   0
      Left            =   300
      TabIndex        =   9
      Top             =   525
      Width           =   915
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Username"
      Height          =   240
      Index           =   1
      Left            =   225
      TabIndex        =   8
      Top             =   150
      Width           =   990
   End
End
Attribute VB_Name = "frmUserRecAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public srcText              As TextBox 'Used in pop-up mode

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset

Private Sub DisplayForEditing()
    On Error GoTo err
    
    With rs
        txtEntry(0).Text = .Fields("UserID")
        txtEntry(1).Text = Enc.DecryptString(.Fields("Password"))
        
        txtEntry(2).Text = .Fields("CompleteName")
        Check1.Value = changeYNValue(.Fields("Admin"))
    End With
    
    Exit Sub
err:
        If err.Number = 94 Then Resume Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetFields()
    clearText Me
    Check1.Value = 0
    txtEntry(0).SetFocus
End Sub

Private Sub cmdSave_Click()
    If is_empty(txtEntry(0), True) = True Then Exit Sub
    If is_empty(txtEntry(1), True) = True Then Exit Sub
    If is_empty(txtEntry(2), True) = True Then Exit Sub
    
    If State = adStateAddMode Then
        rs.AddNew
        rs.Fields("PK") = PK
        rs.Fields("DateAdded") = Now
        rs.Fields("AddedByFK") = CurrUser.USER_PK
    Else
        rs.Fields("DateModified") = Now
        rs.Fields("LastUserFK") = CurrUser.USER_PK
    End If
    'Phill 2:12
    With rs
        .Fields("UserID") = txtEntry(0).Text
        .Fields("Password") = Enc.EncryptString(txtEntry(1).Text)
        .Fields("CompleteName") = txtEntry(2).Text
        .Fields("Admin") = changeYNValue(Check1.Value)
        .Update
    End With
    
    HaveAction = True
    
    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            ResetFields
            PK = getIndex("tbl_SM_Users")
         Else
            Unload Me
        End If
    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If
End Sub

Private Sub cmdUsrHistory_Click()
    On Error Resume Next
    Dim tDate1 As String
    Dim tDate2 As String
    Dim tUser1 As String
    Dim tUser2 As String
    
    tDate1 = Format$(rs.Fields("DateAdded"), "MMM-dd-yyyy HH:MM AMPM")
    tDate2 = Format$(rs.Fields("DateModified"), "MMM-dd-yyyy HH:MM AMPM")
    
    tUser1 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & rs.Fields("AddedByFK"), "CompleteName")
    tUser2 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & rs.Fields("LastUserFK"), "CompleteName")
    
    MsgBox "Date Added: " & tDate1 & vbCrLf & _
           "Added By: " & tUser1 & vbCrLf & _
           "" & vbCrLf & _
           "Last Modified: " & tDate2 & vbCrLf & _
           "Modified By: " & tUser2, vbInformation, "Modification History"
           
    tDate1 = vbNullString
    tDate2 = vbNullString
    tUser1 = vbNullString
    tUser2 = vbNullString
End Sub

Private Sub Form_Load()
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM tbl_SM_Users WHERE PK = " & PK, CN, adOpenStatic, adLockOptimistic
    'Check the form state
    If State = adStateAddMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        PK = getIndex("tbl_SM_Users")
    Else
        Caption = "Edit Entry"
        DisplayForEditing
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or adStateEditMode Then
            frmUserRec.CommandPass 5
        End If
    End If
    
    Set frmUserRecAE = Nothing
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    HLText txtEntry(Index)
End Sub

