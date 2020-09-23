VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPDCManagerAE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   8790
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPDCManagerAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -525
      TabIndex        =   25
      Top             =   3225
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   53
   End
   Begin VB.CheckBox chkCleared 
      Caption         =   "Cleared"
      Height          =   315
      Left            =   5550
      TabIndex        =   10
      Top             =   2025
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   0
      Left            =   1350
      MaxLength       =   100
      TabIndex        =   0
      Tag             =   "Check No"
      Top             =   150
      Width           =   2040
   End
   Begin VB.CommandButton cmdBrowser 
      Height          =   315
      Left            =   3825
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Browse..."
      Top             =   2400
      Width           =   315
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   225
      TabIndex        =   12
      Top             =   3375
      Width           =   1680
   End
   Begin VB.TextBox txtEntry 
      Height          =   1815
      Index           =   6
      Left            =   5550
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Tag             =   "Remarks"
      Top             =   150
      Width           =   3105
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   3
      Left            =   1350
      MaxLength       =   10
      TabIndex        =   5
      Tag             =   "Check Amount"
      Text            =   "0.00"
      Top             =   2025
      Width           =   1665
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1350
      MaxLength       =   100
      TabIndex        =   4
      Tag             =   "Account No"
      Top             =   1650
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1350
      MaxLength       =   100
      TabIndex        =   3
      Tag             =   "Account Name"
      Top             =   1275
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   5
      Left            =   1350
      MaxLength       =   200
      TabIndex        =   8
      Top             =   2775
      Width           =   5190
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   4
      Left            =   1350
      MaxLength       =   100
      TabIndex        =   6
      Tag             =   "Name"
      Top             =   2400
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   7290
      TabIndex        =   14
      Top             =   3375
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   5850
      TabIndex        =   13
      Top             =   3375
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtpIssue 
      Height          =   285
      Left            =   1350
      TabIndex        =   1
      Top             =   525
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   44892163
      CurrentDate     =   38207
   End
   Begin MSComCtl2.DTPicker dtpDue 
      Height          =   285
      Left            =   1350
      TabIndex        =   2
      Top             =   900
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   44892163
      CurrentDate     =   38207
   End
   Begin MSDataListLib.DataCombo dcVan 
      Height          =   315
      Left            =   5550
      TabIndex        =   11
      Top             =   2400
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Van"
      Height          =   240
      Index           =   6
      Left            =   4275
      TabIndex        =   24
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Date Due"
      Height          =   240
      Index           =   12
      Left            =   75
      TabIndex        =   23
      Top             =   900
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Date Issued"
      Height          =   240
      Index           =   11
      Left            =   75
      TabIndex        =   22
      Top             =   525
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Remarks"
      Height          =   240
      Index           =   8
      Left            =   4500
      TabIndex        =   21
      Top             =   150
      Width           =   990
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Check Amount"
      Height          =   240
      Index           =   5
      Left            =   -75
      TabIndex        =   20
      Top             =   2025
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Account No"
      Height          =   240
      Index           =   4
      Left            =   -75
      TabIndex        =   19
      Top             =   1650
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Account Name"
      Height          =   240
      Index           =   3
      Left            =   -75
      TabIndex        =   18
      Top             =   1275
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Address"
      Height          =   240
      Index           =   2
      Left            =   -75
      TabIndex        =   17
      Top             =   2775
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Bank Name"
      Height          =   240
      Index           =   1
      Left            =   75
      TabIndex        =   16
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Check No."
      Height          =   240
      Index           =   0
      Left            =   375
      TabIndex        =   15
      Top             =   150
      Width           =   915
   End
End
Attribute VB_Name = "frmPDCManagerAE"
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
        txtEntry(0).Text = .Fields("CheckNo")
        dtpIssue.Value = .Fields("DateIssued")
        dtpDue.Value = .Fields("DateDue")
        txtEntry(1).Text = .Fields("AccountName")
        txtEntry(2).Text = .Fields("AccountNo")
        txtEntry(3).Text = toMoney(.Fields("CheckAmount"))
        txtEntry(4).Text = .Fields("BankName")
        txtEntry(5).Text = .Fields("Address")
        txtEntry(6).Text = .Fields("Remarks")
        chkCleared.Value = changeYNValue(.Fields("Cleared"))
        dcVan.BoundText = ![VanFK]
    End With
    
    Exit Sub
err:
        If err.Number = 94 Then Resume Next
End Sub

Private Sub cmdBrowser_Click()
    With frmSelectBank
        Set .srcTextBank = txtEntry(4)
        Set .srcTextBAddress = txtEntry(5)
    
        .show vbModal
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetFields()
    clearText Me
    
    dtpIssue.Value = Date
    dtpDue.Value = Date
    chkCleared.Value = 0
    
    txtEntry(0).SetFocus
End Sub

Private Sub GeneratePK()
    PK = getIndex("tbl_AR_PDCManager")
End Sub

Private Sub cmdSave_Click()
    If is_empty(txtEntry(0), True) = True Then Exit Sub
    If is_empty(txtEntry(1), True) = True Then Exit Sub
    If is_empty(txtEntry(2), True) = True Then Exit Sub
    If toNumber(txtEntry(3).Text) < 1 Then
        MsgBox "Check amouny must have a non-zero value.", vbExclamation
        txtEntry(3).SetFocus
        Exit Sub
    End If
    
    If State = adStateAddMode Or State = adStatePopupMode Then
        rs.AddNew
        rs.Fields("PK") = PK
        rs.Fields("DateAdded") = Now
        rs.Fields("AddedByFK") = CurrUser.USER_PK
    Else
        rs.Fields("DateModified") = Now
        rs.Fields("LastUserFK") = CurrUser.USER_PK
    End If
    
    With rs
        .Fields("CheckNo") = txtEntry(0).Text
        .Fields("DateIssued") = dtpIssue.Value
        .Fields("DateDue") = dtpDue.Value
        .Fields("AccountName") = txtEntry(1).Text
        .Fields("AccountNo") = txtEntry(2).Text
        .Fields("CheckAmount") = toNumber(txtEntry(3).Text)
        .Fields("BankName") = txtEntry(4).Text
        .Fields("Address") = txtEntry(5).Text
        .Fields("Remarks") = txtEntry(6).Text
        .Fields("Cleared") = changeYNValue(chkCleared.Value)
        .Fields("VanFK") = dcVan.BoundText
    
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
        MsgBox "New record has been successfully saved.", vbInformation
        Unload Me
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
    
    'Set the controls default property
    dtpIssue.Value = Date
    dtpDue.Value = Date
    
    bind_dc "SELECT * FROM tbl_AR_Van", "VanName", dcVan, "PK", True
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM tbl_AR_PDCManager WHERE PK = " & PK, CN, adOpenStatic, adLockOptimistic
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        If State = adStatePopupMode Then
            Caption = "New PDC Entry"
        Else
            Caption = "Create New Entry"
        End If
        cmdUsrHistory.Enabled = False
        GeneratePK
    Else
        Caption = "Edit Entry"
        DisplayForEditing
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or State = adStateEditMode Then
            frmPDCManager.RefreshRecords
        ElseIf State = adStatePopupMode Then
            If isObjectSet(srcText) = True Then
                srcText.Text = rs![CheckNo]
                srcText.Tag = rs![PK]
            End If
        End If
        MAIN.UpdateInfoMsg
    End If
    
    Set frmPDCManagerAE = Nothing
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    If Index = 6 Then cmdSave.Default = False
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 3 Then KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
    If Index = 6 Then cmdSave.Default = True
End Sub

Private Sub txtEntry_Validate(Index As Integer, Cancel As Boolean)
    If Index = 3 Then txtEntry(Index).Text = toMoney(toNumber(txtEntry(Index).Text))
End Sub
