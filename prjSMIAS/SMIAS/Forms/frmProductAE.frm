VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmProductAE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   465
   ClientWidth     =   8115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProductAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   8115
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   0
      Left            =   1500
      TabIndex        =   0
      Top             =   450
      Width           =   2040
   End
   Begin VB.ComboBox cmbStat 
      Height          =   315
      ItemData        =   "frmProductAE.frx":0A02
      Left            =   6450
      List            =   "frmProductAE.frx":0A0C
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   3600
      Width           =   1290
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   12
      Left            =   6450
      TabIndex        =   14
      Text            =   "0"
      Top             =   2925
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   11
      Left            =   6450
      TabIndex        =   13
      Text            =   "0"
      Top             =   2550
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   10
      Left            =   6450
      TabIndex        =   12
      Text            =   "0"
      Top             =   2175
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   9
      Left            =   6450
      TabIndex        =   11
      Text            =   "0.00"
      Top             =   1500
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   8
      Left            =   6450
      TabIndex        =   10
      Text            =   "0.00"
      Top             =   825
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   7
      Left            =   6450
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   450
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   6
      Left            =   1500
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   4050
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   5
      Left            =   1500
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   3675
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   4
      Left            =   1500
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   3300
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   3
      Left            =   1500
      TabIndex        =   5
      Top             =   2625
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1500
      TabIndex        =   4
      Top             =   2250
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1500
      TabIndex        =   1
      Top             =   825
      Width           =   2490
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   225
      TabIndex        =   16
      Top             =   4575
      Width           =   1680
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   6615
      TabIndex        =   18
      Top             =   4575
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   5175
      TabIndex        =   17
      Top             =   4575
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   1500
      TabIndex        =   2
      Top             =   1200
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   1500
      TabIndex        =   3
      Top             =   1575
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -225
      TabIndex        =   41
      Top             =   4425
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   53
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   225
      TabIndex        =   40
      Top             =   150
      Width           =   3015
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Main Product"
      Height          =   240
      Index           =   15
      Left            =   5325
      TabIndex        =   39
      Top             =   3600
      Width           =   1065
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Box/Case"
      Height          =   240
      Index           =   14
      Left            =   5175
      TabIndex        =   38
      Top             =   2925
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Pieces/Case"
      Height          =   240
      Index           =   13
      Left            =   5175
      TabIndex        =   37
      Top             =   2550
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Pieces/Box"
      Height          =   240
      Index           =   12
      Left            =   5175
      TabIndex        =   36
      Top             =   2175
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty Setup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   5250
      TabIndex        =   35
      Top             =   1875
      Width           =   3015
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Unit Cost (Each)"
      Height          =   240
      Index           =   11
      Left            =   5175
      TabIndex        =   34
      Top             =   1500
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "SRP Price/Pack"
      Height          =   240
      Index           =   10
      Left            =   5175
      TabIndex        =   33
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Price"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   5250
      TabIndex        =   32
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "SRP Price/Pcs."
      Height          =   240
      Index           =   9
      Left            =   5175
      TabIndex        =   31
      Top             =   450
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Suggested Retail Price"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   5250
      TabIndex        =   30
      Top             =   150
      Width           =   2265
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Van Price/Boxes"
      Height          =   240
      Index           =   8
      Left            =   225
      TabIndex        =   29
      Top             =   4050
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Van Price/Case"
      Height          =   240
      Index           =   7
      Left            =   225
      TabIndex        =   28
      Top             =   3675
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Van Price/Pcs."
      Height          =   240
      Index           =   6
      Left            =   225
      TabIndex        =   27
      Top             =   3300
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Van Pricing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   225
      TabIndex        =   26
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Packing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   225
      TabIndex        =   25
      Top             =   1950
      Width           =   3015
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Pack"
      Height          =   240
      Index           =   5
      Left            =   225
      TabIndex        =   24
      Top             =   2625
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Size"
      Height          =   240
      Index           =   4
      Left            =   225
      TabIndex        =   23
      Top             =   2250
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Category"
      Height          =   240
      Index           =   3
      Left            =   225
      TabIndex        =   22
      Top             =   1575
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Supplier"
      Height          =   240
      Index           =   2
      Left            =   225
      TabIndex        =   21
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Description"
      Height          =   240
      Index           =   1
      Left            =   225
      TabIndex        =   20
      Top             =   825
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Product Code"
      Height          =   240
      Index           =   0
      Left            =   225
      TabIndex        =   19
      Top             =   450
      Width           =   1215
   End
   Begin VB.Shape shpBar 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   150
      Top             =   1950
      Width           =   3840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   150
      Top             =   3000
      Width           =   2865
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   5175
      Top             =   150
      Width           =   2790
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   5175
      Top             =   1200
      Width           =   2790
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   5175
      Top             =   1875
      Width           =   2790
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   150
      Top             =   150
      Width           =   3915
   End
End
Attribute VB_Name = "frmProductAE"
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
        txtEntry(0).Text = .Fields("ProductCode")
        txtEntry(1).Text = .Fields("Description")
        
        DataCombo1.BoundText = .Fields("SupplierFK")
        DataCombo2.BoundText = .Fields("CategoryFK")
        
        txtEntry(2).Text = .Fields("Size")
        txtEntry(3).Text = .Fields("Pack")
        
        txtEntry(4).Text = toMoney(.Fields("VPP"))
        txtEntry(5).Text = toMoney(.Fields("VPB"))
        txtEntry(6).Text = toMoney(.Fields("VPC"))
        
        txtEntry(7).Text = toMoney(.Fields("SRPP"))
        txtEntry(8).Text = toMoney(.Fields("SRPPack"))
        
        txtEntry(9).Text = toMoney(.Fields("UnitCost"))
        
        txtEntry(10).Text = .Fields("PiecesPerBox")
        txtEntry(11).Text = .Fields("PiecesPerCase")
        txtEntry(12).Text = .Fields("BoxPerCase")
        
        If .Fields("MainProduct") = "N" Then
            cmbStat.ListIndex = 1
        Else
            cmbStat.ListIndex = 0
        End If
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
    cmbStat.ListIndex = 0
    
    DataCombo1.BoundText = RightSplitUF(DataCombo1.Tag)
    DataCombo2.BoundText = RightSplitUF(DataCombo2.Tag)
    
    txtEntry(0).SetFocus
End Sub

Private Sub GeneratePK()
    PK = getIndex("tbl_IC_Products")
End Sub

Private Sub cmdSave_Click()
    If is_empty(txtEntry(0), True) = True Then Exit Sub
    
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
        .Fields("ProductCode") = txtEntry(0).Text
        .Fields("Description") = txtEntry(1).Text
        
        .Fields("SupplierFK") = DataCombo1.BoundText
        .Fields("CategoryFK") = DataCombo2.BoundText
        
        .Fields("Size") = txtEntry(2).Text
        .Fields("Pack") = txtEntry(3).Text
        
        .Fields("VPP") = toNumber(txtEntry(4).Text)
        .Fields("VPB") = toNumber(txtEntry(5).Text)
        .Fields("VPC") = toNumber(txtEntry(6).Text)
        
        .Fields("SRPP") = toNumber(txtEntry(7).Text)
        .Fields("SRPPack") = toNumber(txtEntry(8).Text)
        
        .Fields("UnitCost") = toNumber(txtEntry(9).Text)
        
        .Fields("PiecesPerBox") = toNumber(txtEntry(10).Text)
        .Fields("PiecesPerCase") = toNumber(txtEntry(11).Text)
        .Fields("BoxPerCase") = toNumber(txtEntry(12).Text)
        
        .Fields("MainProduct") = cmbStat.Text
        
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
    rs.Open "SELECT * FROM tbl_IC_Products WHERE PK = " & PK, CN, adOpenStatic, adLockOptimistic
    
    
    'Bind the data combo
    bind_dc "SELECT * FROM tbl_AP_Supplier", "Name", DataCombo1, "PK", True
    bind_dc "SELECT * FROM tbl_IC_Category", "CategoryName", DataCombo2, "PK", True
    
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        GeneratePK
        cmbStat.ListIndex = 0
    Else
        Caption = "Edit Entry"
        DisplayForEditing
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or adStateEditMode Then frmProduct.RefreshRecords
        MAIN.UpdateInfoMsg
    End If
    
    Set frmProductAE = Nothing
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index > 3 And Index < 13 Then
        KeyAscii = isNumber(KeyAscii)
    End If
End Sub

Private Sub txtEntry_Validate(Index As Integer, Cancel As Boolean)
    If Index > 3 And Index < 10 Then
        txtEntry(Index).Text = toMoney(toNumber(txtEntry(Index).Text))
    ElseIf Index > 9 And Index < 13 Then
        txtEntry(Index).Text = toNumber(txtEntry(Index).Text)
    End If
End Sub
