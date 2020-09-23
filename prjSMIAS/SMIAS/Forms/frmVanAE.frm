VERSION 5.00
Begin VB.Form frmVanAE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   5190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVanAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5190
   StartUpPosition =   2  'CenterScreen
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -1500
      TabIndex        =   9
      Top             =   3000
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   53
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1575
      MaxLength       =   20
      TabIndex        =   2
      Tag             =   "Category Name"
      Top             =   2475
      Width           =   2040
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   150
      TabIndex        =   3
      Top             =   3225
      Width           =   1680
   End
   Begin VB.TextBox txtEntry 
      Height          =   1815
      Index           =   1
      Left            =   1575
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Tag             =   "Remarks"
      Top             =   525
      Width           =   3405
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   0
      Left            =   1575
      MaxLength       =   100
      TabIndex        =   0
      Tag             =   "Category Name"
      Top             =   150
      Width           =   3390
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3690
      TabIndex        =   5
      Top             =   3225
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   2250
      TabIndex        =   4
      Top             =   3225
      Width           =   1335
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Plate No"
      Height          =   240
      Index           =   0
      Left            =   300
      TabIndex        =   8
      Top             =   2475
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Description"
      Height          =   240
      Index           =   8
      Left            =   225
      TabIndex        =   7
      Top             =   525
      Width           =   1290
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Category Name"
      Height          =   240
      Index           =   1
      Left            =   300
      TabIndex        =   6
      Top             =   150
      Width           =   1215
   End
End
Attribute VB_Name = "frmVanAE"
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
        txtEntry(0).Text = .Fields("VanName")
        txtEntry(1).Text = .Fields("Description")
        txtEntry(2).Text = .Fields("PlateNo")
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
    txtEntry(0).SetFocus
End Sub

Private Sub cmdSave_Click()
    If is_empty(txtEntry(0), True) = True Then Exit Sub
    
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
        .Fields("VanName") = txtEntry(0).Text
        .Fields("Description") = txtEntry(1).Text
        .Fields("PlateNo") = txtEntry(2).Text
        .Update
    End With
    
    HaveAction = True
    
    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            ResetFields
            PK = getIndex("tbl_AR_Van")
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
    rs.Open "SELECT * FROM tbl_AR_Van WHERE PK = " & PK, CN, adOpenStatic, adLockOptimistic
    'Check the form state
    If State = adStateAddMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        PK = getIndex("tbl_AR_Van")
    Else
        Caption = "Edit Entry"
        DisplayForEditing
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or adStateEditMode Then
            frmVan.RefreshRecords
        End If
    End If
    
    Set frmVanAE = Nothing
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    If Index = 8 Then cmdSave.Default = False
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
    If Index = 8 Then cmdSave.Default = True
End Sub
