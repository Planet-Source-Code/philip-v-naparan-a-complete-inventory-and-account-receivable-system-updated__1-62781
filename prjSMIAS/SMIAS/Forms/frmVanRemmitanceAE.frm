VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmVanRemmitanceAE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVanRemmitanceAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtVan 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5025
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   225
      Width           =   2550
   End
   Begin VB.TextBox txtRNo 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   225
      Width           =   2115
   End
   Begin VB.TextBox txtReadOnly 
      BackColor       =   &H00E6FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   9
      Left            =   6075
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "0.00"
      Top             =   3150
      Width           =   1500
   End
   Begin VB.TextBox txtReadOnly 
      BackColor       =   &H00E6FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Index           =   8
      Left            =   6075
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "0.00"
      Top             =   2775
      Width           =   1500
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   6075
      TabIndex        =   14
      Text            =   "0.00"
      Top             =   2400
      Width           =   1515
   End
   Begin VB.TextBox txtReadOnly 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   7
      Left            =   6075
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "0.00"
      Top             =   1875
      Width           =   1515
   End
   Begin VB.TextBox txtReadOnly 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   6
      Left            =   6075
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "0.00"
      Top             =   1500
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   0
      Left            =   1575
      TabIndex        =   10
      Text            =   "0.00"
      Top             =   4125
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Height          =   990
      Index           =   2
      Left            =   1575
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Tag             =   "Remarks"
      Top             =   4500
      Width           =   6030
   End
   Begin VB.TextBox txtReadOnly 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   3300
      Width           =   1515
   End
   Begin VB.TextBox txtReadOnly 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   5
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   3675
      Width           =   1515
   End
   Begin VB.TextBox txtReadOnly 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   2250
      Width           =   1515
   End
   Begin VB.TextBox txtReadOnly 
      BackColor       =   &H00E6FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   2625
      Width           =   1515
   End
   Begin VB.TextBox txtReadOnly 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   1500
      Width           =   1515
   End
   Begin VB.TextBox txtReadOnly 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   1875
      Width           =   1515
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   300
      TabIndex        =   17
      Top             =   5775
      Width           =   1755
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   6315
      TabIndex        =   19
      Top             =   5775
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   4875
      TabIndex        =   18
      Top             =   5775
      Width           =   1335
   End
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -675
      TabIndex        =   20
      Top             =   5625
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   53
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   285
      Left            =   1575
      TabIndex        =   1
      Top             =   600
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   19595267
      CurrentDate     =   38207
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   2100
   End
   Begin SMIAS.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   150
      TabIndex        =   39
      Top             =   975
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   53
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Van"
      Height          =   240
      Index           =   0
      Left            =   3750
      TabIndex        =   40
      Top             =   225
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Remmitance No"
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
      Index           =   4
      Left            =   -75
      TabIndex        =   38
      Top             =   225
      Width           =   1590
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Remmitance"
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
      Left            =   4575
      TabIndex        =   37
      Top             =   1200
      Width           =   1590
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Over"
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   4500
      TabIndex        =   36
      Top             =   3150
      Width           =   1515
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Short"
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   4500
      TabIndex        =   35
      Top             =   2775
      Width           =   1515
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Amount Remmited"
      Height          =   240
      Index           =   11
      Left            =   4275
      TabIndex        =   34
      Top             =   2400
      Width           =   1740
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Net Cash"
      Height          =   240
      Index           =   10
      Left            =   4800
      TabIndex        =   33
      Top             =   1875
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Cash"
      Height          =   240
      Index           =   9
      Left            =   4800
      TabIndex        =   32
      Top             =   1500
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Less"
      Height          =   240
      Index           =   8
      Left            =   300
      TabIndex        =   31
      Top             =   4125
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Remarks"
      Height          =   240
      Index           =   7
      Left            =   525
      TabIndex        =   30
      Top             =   4500
      Width           =   990
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "PDC Collection"
      Height          =   240
      Index           =   6
      Left            =   300
      TabIndex        =   29
      Top             =   3300
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Cash Collection"
      Height          =   240
      Index           =   5
      Left            =   300
      TabIndex        =   28
      Top             =   3675
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Collections"
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
      Left            =   375
      TabIndex        =   27
      Top             =   3000
      Width           =   2715
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Charge Account"
      Height          =   240
      Index           =   3
      Left            =   300
      TabIndex        =   26
      Top             =   2250
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Sales"
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
      Left            =   300
      TabIndex        =   25
      Top             =   2625
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
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
      Left            =   375
      TabIndex        =   24
      Top             =   1200
      Width           =   3015
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Cash Sales"
      Height          =   240
      Index           =   12
      Left            =   300
      TabIndex        =   23
      Top             =   1500
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "PDC Sales"
      Height          =   240
      Index           =   13
      Left            =   300
      TabIndex        =   22
      Top             =   1875
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Date Remmited"
      Height          =   240
      Index           =   1
      Left            =   300
      TabIndex        =   21
      Top             =   600
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   300
      Top             =   1200
      Width           =   2790
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   300
      Top             =   3000
      Width           =   2790
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   4500
      Top             =   1200
      Width           =   3090
   End
End
Attribute VB_Name = "frmVanRemmitanceAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public LLFK                 As Long 'Last loading FK
Public CloseMe              As Boolean

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset

Private Sub DisplayForViewing()
    On Error GoTo err
    
    With rs
        txtRNo.Text = .Fields("RemmitanceNo")
        txtDate.Text = Format$(.Fields("DateRemmited"), "MMM-dd-yyyy")
        txtVan.Text = .Fields("VanName")
        
        txtReadOnly(0).Tag = toNumber(.Fields("CashSales"))
        txtReadOnly(1).Tag = toNumber(.Fields("PDCSales"))
        txtReadOnly(2).Tag = toNumber(.Fields("ChargeAccount"))
        txtReadOnly(5).Tag = toNumber(.Fields("CashCollection"))
        txtReadOnly(4).Tag = toNumber(.Fields("PDCCollection"))
        txtEntry(0).Text = toMoney(toNumber(.Fields("Less")))
        txtEntry(2).Text = .Fields("Remarks")
        txtEntry(1).Text = toMoney(toNumber(.Fields("CashRemitted")))

        .Update
    End With
        
    txtReadOnly(0).Text = toMoney(toNumber(txtReadOnly(0).Tag))
    txtReadOnly(1).Text = toMoney(toNumber(txtReadOnly(1).Tag))
    txtReadOnly(2).Text = toMoney(toNumber(txtReadOnly(2).Tag))
    
    txtReadOnly(3).Tag = toNumber(txtReadOnly(0).Tag) + toNumber(txtReadOnly(1).Tag) + toNumber(txtReadOnly(2).Tag)
    txtReadOnly(3).Text = toMoney(toNumber(txtReadOnly(3).Tag))
    
    txtReadOnly(4).Text = toMoney(toNumber(txtReadOnly(4).Tag))
    txtReadOnly(5).Text = toMoney(toNumber(txtReadOnly(5).Tag))
    
    txtReadOnly(6).Tag = toNumber(txtReadOnly(0).Tag) + toNumber(txtReadOnly(1).Tag) + toNumber(txtReadOnly(5).Tag)
    txtReadOnly(6).Text = toMoney(toNumber(txtReadOnly(6).Tag))
    
    txtEntry(0).Locked = True
    txtEntry(1).Locked = True
    txtEntry(2).Locked = True
    dtpDate.Visible = False
    txtDate.Visible = True
    
    Exit Sub
err:
        If err.Number = 94 Then Resume Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If MsgBox("This save the record.Do you want to proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    On Error GoTo err
    
    CN.BeginTrans
    
    rs.AddNew
    rs.Fields("PK") = PK
    rs.Fields("RemmitanceNo") = txtRNo.Text
    rs.Fields("DateAdded") = Now
    rs.Fields("AddedByFK") = CurrUser.USER_PK
    
    With rs
        .Fields("Date") = dtpDate.Value
        .Fields("VanFK") = toNumber(txtVan.Tag)
        .Fields("LLFK") = LLFK
        .Fields("CashSales") = toNumber(txtReadOnly(0).Text)
        .Fields("PDCSales") = toNumber(txtReadOnly(1).Text)
        .Fields("ChargeAccount") = toNumber(txtReadOnly(2).Text)
        .Fields("CashCollection") = toNumber(txtReadOnly(5).Text)
        .Fields("PDCCollection") = toNumber(txtReadOnly(4).Text)
        .Fields("Less") = toNumber(txtEntry(0).Text)
        .Fields("Remarks") = txtEntry(2).Text
        .Fields("CashRemitted") = toNumber(txtEntry(1).Text)
        .Fields("Short") = toNumber(txtReadOnly(8).Text)
        .Fields("Over") = toNumber(txtReadOnly(9).Text)

        .Update
    End With
    
    'Lock loading and all it's transactions
    ChangeValue CN, "tbl_IC_Loading", "Lock", "Y", , "WHERE PK=" & LLFK
    ChangeValue CN, "tbl_AR_Invoice", "Lock", "Y", , "WHERE LastLoadingFK=" & LLFK
    ChangeValue CN, "tbl_AR_VanCollection", "Lock", "Y", , "WHERE LLFK=" & LLFK
    ChangeValue CN, "tbl_IC_VanInv", "Lock", "Y", , "WHERE LLFK=" & LLFK
    
    CN.CommitTrans
    
    HaveAction = True
    
    If State = adStateAddMode Or State = adStatePopupMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        Unload Me
    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If
    
    Exit Sub
err:
    CN.RollbackTrans
    prompt_err err, Me.Name, "cmdSave_Click"
End Sub

Private Sub cmdUsrHistory_Click()
    On Error Resume Next
    Dim tDate1 As String
    Dim tUser1 As String
    
    tDate1 = Format$(rs.Fields("DateAdded"), "MMM-dd-yyyy HH:MM AMPM")
    
    tUser1 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & rs.Fields("AddedByFK"), "CompleteName")
    
    MsgBox "Date Added: " & tDate1 & vbCrLf & _
           "Added By: " & tUser1 & vbCrLf & _
           "" & vbCrLf & _
           "Last Modified: n/a" & vbCrLf & _
           "Modified By: n/a", vbInformation, "Modification History"
           
    tDate1 = vbNullString
    tUser1 = vbNullString
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If CloseMe = True Then Unload Me: Exit Sub
    txtRNo.SetFocus
End Sub

Private Sub Form_Load()
    dtpDate.Value = Date
         
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        frmVanRemmitanceAEPickFrom.show vbModal
        GetSalesAndCol
        
        'Set the recordset
        rs.Open "SELECT * FROM tbl_AR_VanRemmitance WHERE PK=" & PK, CN, adOpenStatic, adLockOptimistic
        
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        GeneratePK
    Else
        Screen.MousePointer = vbHourglass
        'Set the recordset
        rs.Open "SELECT * FROM qry_AR_VanRemmitance WHERE PK=" & PK, CN, adOpenStatic, adLockOptimistic

        Caption = "View Record"
        cmdCancel.Caption = "Close"
        cmdUsrHistory.Enabled = True
        DisplayForViewing
        
        MsgBox "This is use for viewing the record only." & vbCrLf & _
               "You cannot perform any changes in this form." & vbCrLf & vbCrLf & _
               "Note:If you have mistake in adding this record then " & vbCrLf & _
               "void this record and re-enter.", vbExclamation
               
        Screen.MousePointer = vbDefault
    End If

End Sub

'Procedure used to generate PK
Private Sub GeneratePK()
    PK = getIndex("tbl_AR_VanRemmitance")
    txtRNo.Text = GenerateID(PK, Format$(Date, "yyyy") & Format$(Date, "mm") & Format$(Date, "dd") & "-", "0")
End Sub

Private Sub GetSalesAndCol()
    'Get the sales in the selected loading
    txtReadOnly(0).Tag = getValueAt("SELECT * FROM qry_AR_SalesByLoading WHERE((PaymentType='Cash') AND (LLFK=" & LLFK & "))", "AmountPaid")
    txtReadOnly(1).Tag = getValueAt("SELECT * FROM qry_AR_SalesByLoading WHERE(((PaymentType='On Date Check') OR (PaymentType='Post Dated Check')) AND (LLFK=" & LLFK & "))", "AmountPaid")
    txtReadOnly(2).Tag = getValueAt("SELECT * FROM qry_AR_BalanceByLoading WHERE LLFK=" & LLFK, "Balance")
    
    txtReadOnly(0).Text = toMoney(toNumber(txtReadOnly(0).Tag))
    txtReadOnly(1).Text = toMoney(toNumber(txtReadOnly(1).Tag))
    txtReadOnly(2).Text = toMoney(toNumber(txtReadOnly(2).Tag))
    
    txtReadOnly(3).Tag = toNumber(txtReadOnly(0).Tag) + toNumber(txtReadOnly(1).Tag) + toNumber(txtReadOnly(2).Tag)
    txtReadOnly(3).Text = toMoney(toNumber(txtReadOnly(3).Tag))
    'Get the collection in the selected loading
    txtReadOnly(4).Tag = getValueAt("SELECT * FROM qry_AR_CollectionByLoading WHERE(((PaymentType='On Date Check') OR (PaymentType='Post Dated Check')) AND (LLFK=" & LLFK & "))", "Collection")
    txtReadOnly(5).Tag = getValueAt("SELECT * FROM qry_AR_CollectionByLoading WHERE((PaymentType='Cash') AND (LLFK=" & LLFK & "))", "Collection")
    
    txtReadOnly(4).Text = toMoney(toNumber(txtReadOnly(4).Tag))
    txtReadOnly(5).Text = toMoney(toNumber(txtReadOnly(5).Tag))
    
    'Display total cash
    txtReadOnly(6).Tag = toNumber(txtReadOnly(0).Tag) + toNumber(txtReadOnly(5).Tag)
    txtReadOnly(6).Text = toMoney(toNumber(txtReadOnly(6).Tag))
    
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        frmVanRemmitance.RefreshRecords
        MAIN.UpdateInfoMsg
    End If
    
    Set frmVanRemmitanceAE = Nothing
End Sub

Private Sub txtDate_GotFocus()
    HLText txtDate
End Sub

Private Sub txtEntry_Change(Index As Integer)
    If Index = 0 Then
        txtReadOnly_Change 6
    ElseIf Index = 1 Then
        txtReadOnly(8).Text = toMoney(toNumber(toNumber(txtReadOnly(7).Text) - toNumber(txtEntry(1).Text), True))
        txtReadOnly(9).Text = toMoney(toNumber(toNumber(txtEntry(1).Text) - toNumber(txtReadOnly(7).Text), True))
    End If
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    If Index = 2 Then cmdSave.Default = False
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index < 2 Then
        KeyAscii = isNumber(KeyAscii)
    End If
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
    If Index = 2 Then cmdSave.Default = True
End Sub

Private Sub txtEntry_Validate(Index As Integer, Cancel As Boolean)
    If Index < 2 Then
        txtEntry(Index).Text = toMoney(toNumber(txtEntry(Index).Text))
    End If
End Sub

Private Sub txtReadOnly_Change(Index As Integer)
    Select Case Index
        Case 6: txtReadOnly(7).Text = toMoney(toNumber(txtReadOnly(6).Tag) - toNumber(txtEntry(0).Text)): txtEntry_Change 1
    End Select
End Sub

Private Sub txtReadOnly_GotFocus(Index As Integer)
    HLText txtReadOnly(Index)
End Sub

Private Sub txtRNo_GotFocus()
    HLText txtRNo
End Sub

Private Sub txtVan_GotFocus()
    HLText txtVan
End Sub
