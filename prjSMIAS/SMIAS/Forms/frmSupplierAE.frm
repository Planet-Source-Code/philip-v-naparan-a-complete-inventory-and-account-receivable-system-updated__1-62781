VERSION 5.00
Begin VB.Form frmSupplierAE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   8265
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSupplierAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBrowser 
      Height          =   315
      Left            =   2700
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Browse..."
      Top             =   2025
      Width           =   315
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   150
      TabIndex        =   10
      Top             =   3375
      Width           =   1680
   End
   Begin VB.TextBox txtEntry 
      Height          =   1815
      Index           =   8
      Left            =   5025
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Tag             =   "Remarks"
      Top             =   150
      Width           =   3105
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   6
      Left            =   1350
      MaxLength       =   100
      TabIndex        =   7
      Top             =   2400
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   7
      Left            =   1350
      MaxLength       =   20
      TabIndex        =   8
      Top             =   2775
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   5
      Left            =   1350
      MaxLength       =   10
      TabIndex        =   5
      Top             =   2025
      Width           =   1290
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   4
      Left            =   1350
      MaxLength       =   100
      TabIndex        =   4
      Top             =   1650
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   3
      Left            =   1350
      MaxLength       =   100
      TabIndex        =   3
      Top             =   1275
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1350
      MaxLength       =   200
      TabIndex        =   2
      Top             =   900
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1350
      MaxLength       =   100
      TabIndex        =   1
      Tag             =   "Name"
      Top             =   525
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   6765
      TabIndex        =   12
      Top             =   3375
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   5325
      TabIndex        =   11
      Top             =   3375
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1350
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   1965
   End
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -150
      TabIndex        =   22
      Top             =   3225
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   53
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Memo"
      Height          =   240
      Index           =   8
      Left            =   3975
      TabIndex        =   21
      Top             =   150
      Width           =   990
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Contact No"
      Height          =   240
      Index           =   7
      Left            =   -75
      TabIndex        =   20
      Top             =   2775
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Contact Person"
      Height          =   240
      Index           =   6
      Left            =   -75
      TabIndex        =   19
      Top             =   2400
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Zip Code"
      Height          =   240
      Index           =   5
      Left            =   -75
      TabIndex        =   18
      Top             =   2025
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Province"
      Height          =   240
      Index           =   4
      Left            =   -75
      TabIndex        =   17
      Top             =   1650
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "City/Town"
      Height          =   240
      Index           =   3
      Left            =   -75
      TabIndex        =   16
      Top             =   1275
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Address"
      Height          =   240
      Index           =   2
      Left            =   -75
      TabIndex        =   15
      Top             =   900
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      Height          =   240
      Index           =   1
      Left            =   675
      TabIndex        =   14
      Top             =   525
      Width           =   615
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Supplier ID"
      Height          =   240
      Index           =   0
      Left            =   375
      TabIndex        =   13
      Top             =   150
      Width           =   915
   End
End
Attribute VB_Name = "frmSupplierAE"
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
        txtEntry(0).Text = .Fields("SupplierID")
        txtEntry(1).Text = .Fields("Name")
        txtEntry(2).Text = .Fields("Address")
        txtEntry(3).Text = .Fields("CityTown")
        txtEntry(4).Text = .Fields("Province")
        txtEntry(5).Text = .Fields("ZipCode")
        txtEntry(6).Text = .Fields("ContactPerson")
        txtEntry(7).Text = .Fields("ContactNo")
        txtEntry(8).Text = .Fields("Memo")
    End With
    
    Exit Sub
err:
        If err.Number = 94 Then Resume Next
End Sub

Private Sub cmdBrowser_Click()
    With frmSelectZipCode
        Set .txtCity = txtEntry(3)
        Set .txtState = txtEntry(4)
        Set .txtZipCode = txtEntry(5)
    
        .show vbModal
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetFields()
    clearText Me
    txtEntry(1).SetFocus
End Sub

Private Sub GeneratePK()
    PK = getIndex("tbl_AP_Supplier")
    txtEntry(0).Text = GenerateID(PK, "SUP-", "00000")
End Sub

Private Sub cmdSave_Click()
    If is_empty(txtEntry(1), True) = True Then Exit Sub
    
    If State = adStateAddMode Or State = adStatePopupMode Then
        rs.AddNew
        rs.Fields("PK") = PK
        rs.Fields("SupplierID") = txtEntry(0).Text
        rs.Fields("DateAdded") = Now
        rs.Fields("AddedByFK") = CurrUser.USER_PK
    Else
        rs.Fields("DateModified") = Now
        rs.Fields("LastUserFK") = CurrUser.USER_PK
    End If
    
    With rs
        .Fields("Name") = txtEntry(1).Text
        .Fields("Address") = txtEntry(2).Text
        .Fields("CityTown") = txtEntry(3).Text
        .Fields("Province") = txtEntry(4).Text
        .Fields("ZipCode") = txtEntry(5).Text
        .Fields("ContactPerson") = txtEntry(6).Text
        .Fields("ContactNo") = txtEntry(7).Text
        .Fields("Memo") = txtEntry(8).Text
    
        .Update
    End With
    
    HaveAction = True
    
    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            ResetFields
            GeneratePK
         Else
            Unload Me
        End If
    ElseIf State = adStatePopupMode Then
        'POP-UP MODE HERE
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
    'Set the graphics for the controls
    With MAIN
        cmdBrowser.Picture = .i16x16.ListImages(8).Picture
    End With
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM tbl_AP_Supplier WHERE PK = " & PK, CN, adOpenStatic, adLockOptimistic
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        GeneratePK
    Else
        Caption = "Edit Entry"
        DisplayForEditing
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or adStateEditMode Then
            frmSupplier.RefreshRecords
        ElseIf State = adStatePopupMode Then
            srcText.Text = rs![Name]
            srcText.Tag = rs![PK]
        End If
    End If
    
    Set frmSupplierAE = Nothing
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    If Index = 8 Then cmdSave.Default = False
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
    If Index = 8 Then cmdSave.Default = True
End Sub
