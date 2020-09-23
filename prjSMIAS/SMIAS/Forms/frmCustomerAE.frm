VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCustomerAE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   360
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
   Icon            =   "frmCustomerAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPH 
      Caption         =   "Purchase History"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1950
      TabIndex        =   14
      Top             =   3375
      Width           =   1590
   End
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -3000
      TabIndex        =   28
      Top             =   3225
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   53
   End
   Begin VB.ComboBox cmbStat 
      Height          =   315
      ItemData        =   "frmCustomerAE.frx":0A02
      Left            =   5025
      List            =   "frmCustomerAE.frx":0A0C
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   2400
      Width           =   1290
   End
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
      TabIndex        =   13
      Top             =   3375
      Width           =   1680
   End
   Begin VB.ComboBox cmdDisc 
      Height          =   315
      Left            =   5025
      TabIndex        =   9
      Top             =   150
      Width           =   1290
   End
   Begin VB.TextBox txtEntry 
      Height          =   1815
      Index           =   8
      Left            =   5025
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Tag             =   "Remarks"
      Top             =   525
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
      TabIndex        =   16
      Top             =   3375
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   5325
      TabIndex        =   15
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
   Begin MSDataListLib.DataCombo dcVan 
      Height          =   315
      Left            =   5025
      TabIndex        =   12
      Top             =   2775
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
      Index           =   11
      Left            =   3750
      TabIndex        =   29
      Top             =   2775
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Status"
      Height          =   240
      Index           =   10
      Left            =   3900
      TabIndex        =   27
      Top             =   2400
      Width           =   1065
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Discount(%)"
      Height          =   240
      Index           =   9
      Left            =   3900
      TabIndex        =   26
      Top             =   150
      Width           =   1065
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Memo"
      Height          =   240
      Index           =   8
      Left            =   3975
      TabIndex        =   25
      Top             =   525
      Width           =   990
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Contact No"
      Height          =   240
      Index           =   7
      Left            =   -75
      TabIndex        =   24
      Top             =   2775
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Contact Person"
      Height          =   240
      Index           =   6
      Left            =   -75
      TabIndex        =   23
      Top             =   2400
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Zip Code"
      Height          =   240
      Index           =   5
      Left            =   -75
      TabIndex        =   22
      Top             =   2025
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Province"
      Height          =   240
      Index           =   4
      Left            =   -75
      TabIndex        =   21
      Top             =   1650
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "City/Town"
      Height          =   240
      Index           =   3
      Left            =   -75
      TabIndex        =   20
      Top             =   1275
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Address"
      Height          =   240
      Index           =   2
      Left            =   -75
      TabIndex        =   19
      Top             =   900
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      Height          =   240
      Index           =   1
      Left            =   675
      TabIndex        =   18
      Top             =   525
      Width           =   615
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Customer ID"
      Height          =   240
      Index           =   0
      Left            =   375
      TabIndex        =   17
      Top             =   150
      Width           =   915
   End
End
Attribute VB_Name = "frmCustomerAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public srcText              As TextBox 'Used in pop-up mode
Public srcTextAdd           As TextBox 'Used in pop-up mode -> Display the customer address
Public srcTextCP            As TextBox 'Used in pop-up mode -> Display the customer contact person
Public srcTextDisc          As Object  'Used in pop-up mode -> Display the customer Discount (can be combo or textbox)

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset

Private Sub DisplayForEditing()
    On Error GoTo err
    
    With rs
        txtEntry(0).Text = .Fields("CustomerID")
        txtEntry(1).Text = .Fields("Name")
        txtEntry(2).Text = .Fields("Address")
        txtEntry(3).Text = .Fields("CityTown")
        txtEntry(4).Text = .Fields("Province")
        txtEntry(5).Text = .Fields("ZipCode")
        txtEntry(6).Text = .Fields("ContactPerson")
        txtEntry(7).Text = .Fields("ContactNo")
        cmdDisc.Text = .Fields("Discount")
        txtEntry(8).Text = .Fields("Memo")
        If .Fields("Status") = "Old" Then
            cmbStat.ListIndex = 1
        Else
            cmbStat.ListIndex = 0
        End If
        dcVan.BoundText = ![VanFK]
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
    cmbStat.ListIndex = 0
    
    cmdDisc.Text = ""
    
    txtEntry(1).SetFocus
End Sub

Private Sub GeneratePK()
    PK = getIndex("tbl_AR_Customer")
    txtEntry(0).Text = GenerateID(PK, "CUS-", "00000")
End Sub

Private Sub cmdPH_Click()
    frmInvoiceViewer.CUS_PK = PK
    frmInvoiceViewer.Caption = "Purchase History Viewer"
    frmInvoiceViewer.lblTitle.Caption = "Purchase History Viewer"
    frmInvoiceViewer.show vbModal
End Sub

Private Sub cmdSave_Click()
    If is_empty(txtEntry(1), True) = True Then Exit Sub
    If toNumber(cmdDisc) > 99 Then MsgBox "Field value must be 0 to 99 only.", vbExclamation: cmdDisc.SetFocus: Exit Sub
    
    If State = adStateAddMode Or State = adStatePopupMode Then
        rs.AddNew
        rs.Fields("PK") = PK
        rs.Fields("CustomerID") = txtEntry(0).Text
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
        .Fields("Discount") = toNumber(cmdDisc)
        .Fields("Memo") = txtEntry(8).Text
        .Fields("Status") = cmbStat.Text
        .Fields("VanFK") = dcVan.BoundText
        
        ![DisplayAddr] = txtEntry(2).Text
        If txtEntry(3).Text <> "" Then ![DisplayAddr] = ![DisplayAddr] & "," & txtEntry(3).Text
        If txtEntry(4).Text <> "" Then ![DisplayAddr] = ![DisplayAddr] & "," & txtEntry(4).Text
        If txtEntry(5).Text <> "" Then ![DisplayAddr] = ![DisplayAddr] & "," & txtEntry(5).Text
    
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
    
    'Fill the discount combo
    cmdDisc.AddItem "0.01"
    cmdDisc.AddItem "0.02"
    cmdDisc.AddItem "0.03"
    cmdDisc.AddItem "0.04"
    cmdDisc.AddItem "0.05"
    cmdDisc.AddItem "0.06"
    cmdDisc.AddItem "0.07"
    cmdDisc.AddItem "0.08"
    cmdDisc.AddItem "0.09"
    cmdDisc.AddItem "0.1"
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM tbl_AR_Customer WHERE PK = " & PK, CN, adOpenStatic, adLockOptimistic
    
    bind_dc "SELECT * FROM tbl_AR_Van", "VanName", dcVan, "PK", True
    
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        GeneratePK
        cmbStat.ListIndex = 0
    Else
        Caption = "Edit Entry"
        DisplayForEditing
        cmdPH.Enabled = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or State = adStateEditMode Then
            frmCustomer.RefreshRecords
        ElseIf State = adStatePopupMode Then
            srcText.Text = txtEntry(0).Text
            srcText.Tag = PK
            On Error Resume Next
            srcTextAdd.Text = rs![DisplayAddr]
            srcTextCP.Text = txtEntry(6).Text
            srcTextDisc.Text = toNumber(cmdDisc.Text)
        End If
        MAIN.UpdateInfoMsg
    End If
    
    Set frmCustomerAE = Nothing
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    If Index = 8 Then cmdSave.Default = False
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
    If Index = 8 Then cmdSave.Default = True
End Sub
