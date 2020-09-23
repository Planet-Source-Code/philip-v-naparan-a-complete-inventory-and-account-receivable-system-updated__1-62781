VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmVanCollectionAE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVanCollectionAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin SMIAS.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   225
      TabIndex        =   38
      Top             =   2850
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   53
   End
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -75
      TabIndex        =   15
      Top             =   6975
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   53
   End
   Begin VB.TextBox txtEntry 
      Height          =   990
      Index           =   8
      Left            =   225
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Tag             =   "Remarks"
      Top             =   5775
      Width           =   5805
   End
   Begin VB.TextBox txtTA 
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
      Left            =   9450
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "0.00"
      Top             =   5550
      Width           =   1500
   End
   Begin VB.PictureBox picCusInfo 
      BorderStyle     =   0  'None
      Height          =   1740
      Left            =   75
      ScaleHeight     =   1740
      ScaleWidth      =   10965
      TabIndex        =   27
      Top             =   1050
      Width           =   10965
      Begin VB.TextBox txtInv 
         Height          =   315
         Left            =   1350
         TabIndex        =   4
         Top             =   675
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.CheckBox chkOldInv 
         Caption         =   "The invoice already encoded."
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
         Left            =   1350
         TabIndex        =   7
         Top             =   375
         Value           =   1  'Checked
         Width           =   3315
      End
      Begin VB.CommandButton btnCollect 
         Caption         =   "Add To Collection"
         Enabled         =   0   'False
         Height          =   315
         Left            =   9225
         TabIndex        =   14
         Top             =   1425
         Width           =   1635
      End
      Begin VB.TextBox txtRem 
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6375
         TabIndex        =   13
         Top             =   1050
         Width           =   4500
      End
      Begin VB.ComboBox cbPT 
         Height          =   315
         ItemData        =   "frmVanCollectionAE.frx":038A
         Left            =   1350
         List            =   "frmVanCollectionAE.frx":0397
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1080
         Width           =   2565
      End
      Begin VB.TextBox txtBal 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6375
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0.00"
         Top             =   375
         Width           =   1875
      End
      Begin VB.TextBox txtPayment 
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1350
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   1425
         Width           =   1500
      End
      Begin VB.TextBox txtCusAdd 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6375
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   0
         Width           =   3750
      End
      Begin ctrlNSDataCombo.NSDataCombo nsdInvoice 
         Height          =   315
         Left            =   1350
         TabIndex        =   8
         Top             =   675
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtColDate 
         Height          =   285
         Left            =   1350
         TabIndex        =   6
         Top             =   0
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMM-dd-yyyy"
         Format          =   19660803
         CurrentDate     =   38207
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Date"
         Height          =   240
         Index           =   3
         Left            =   75
         TabIndex        =   40
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks"
         ForeColor       =   &H0000011D&
         Height          =   240
         Left            =   4275
         TabIndex        =   36
         Top             =   1050
         Width           =   2040
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Payment Type"
         Height          =   240
         Index           =   12
         Left            =   -900
         TabIndex        =   35
         Top             =   1080
         Width           =   2190
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Balance"
         Height          =   240
         Index           =   6
         Left            =   4125
         TabIndex        =   34
         Top             =   375
         Width           =   2190
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Payment"
         ForeColor       =   &H0000011D&
         Height          =   240
         Left            =   -750
         TabIndex        =   33
         Top             =   1425
         Width           =   2040
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer Name"
         Height          =   240
         Index           =   5
         Left            =   4125
         TabIndex        =   29
         Top             =   0
         Width           =   2190
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Invoice No"
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
         Left            =   -600
         TabIndex        =   28
         Top             =   675
         Width           =   1890
      End
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   300
      Picture         =   "frmVanCollectionAE.frx":03C2
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Remove"
      Top             =   3375
      Visible         =   0   'False
      Width           =   275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   8175
      TabIndex        =   21
      Top             =   7125
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   9615
      TabIndex        =   22
      Top             =   7125
      Width           =   1335
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   225
      TabIndex        =   20
      Top             =   7125
      Width           =   1755
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1425
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   2490
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2190
      Left            =   225
      TabIndex        =   16
      Top             =   3225
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   3863
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   275
      ForeColorFixed  =   -2147483640
      BackColorSel    =   1091552
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      GridColorUnpopulated=   -2147483633
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   285
      Left            =   1425
      TabIndex        =   1
      Top             =   525
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   19660803
      CurrentDate     =   38207
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1425
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   525
      Visible         =   0   'False
      Width           =   2460
   End
   Begin MSDataListLib.DataCombo dcSalesman 
      Height          =   315
      Left            =   7875
      TabIndex        =   5
      Top             =   450
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin VB.TextBox txtNCus 
      Height          =   210
      Left            =   6450
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4725
      Visible         =   0   'False
      Width           =   75
   End
   Begin MSDataListLib.DataCombo dcVan 
      Height          =   315
      Left            =   7875
      TabIndex        =   3
      Top             =   75
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin SMIAS.ctrlLiner ctrlLiner3 
      Height          =   30
      Left            =   225
      TabIndex        =   39
      Top             =   900
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   53
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Van"
      Height          =   240
      Index           =   7
      Left            =   6600
      TabIndex        =   37
      Top             =   75
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Remarks"
      Height          =   240
      Index           =   4
      Left            =   -150
      TabIndex        =   32
      Top             =   5550
      Width           =   990
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Collection"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   7350
      TabIndex        =   30
      Top             =   5550
      Width           =   2040
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Salesman"
      Height          =   240
      Index           =   18
      Left            =   6600
      TabIndex        =   26
      Top             =   450
      Width           =   1215
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Collection"
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
      Left            =   300
      TabIndex        =   25
      Top             =   2925
      Width           =   4365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Date"
      Height          =   240
      Index           =   1
      Left            =   150
      TabIndex        =   24
      Top             =   525
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Collection No"
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
      Index           =   0
      Left            =   150
      TabIndex        =   23
      Top             =   150
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   225
      Top             =   2925
      Width           =   10740
   End
End
Attribute VB_Name = "frmVanCollectionAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public LLFK                 As Long 'Last loading FK
Public CloseMe              As Boolean

Dim cCAmount                As Currency 'Current Collection Amount
Dim cCRowCount              As Integer

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset 'Main recordset for Invoice


Private Sub btnCollect_Click()
    If toNumber(txtPayment.Text) < 0 Then
        MsgBox "Please enter a valid payment.", vbExclamation
        txtPayment.SetFocus
        Exit Sub
    End If

    Dim CurrRow As Integer

    If chkOldInv.Value = 1 Then
        CurrRow = getFlexPos(Grid, 8, nsdInvoice.BoundText)
    Else
        CurrRow = -1
    End If

    'Add to grid
    With Grid
        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 8) = "" And .TextMatrix(1, 5) = "" Then
                .TextMatrix(1, 1) = Format$(dtColDate.Value, "MMM-dd-yyyy")
                If chkOldInv.Value = 1 Then
                    .TextMatrix(1, 2) = nsdInvoice.Text
                    .TextMatrix(1, 8) = nsdInvoice.BoundText
                Else
                    .TextMatrix(1, 2) = txtInv.Text
                End If
                .TextMatrix(1, 3) = txtCusAdd.Text
                .TextMatrix(1, 4) = cbPT.Text
                .TextMatrix(1, 5) = txtPayment.Text
                .TextMatrix(1, 6) = txtBal.Text
                .TextMatrix(1, 7) = txtRem.Text
            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = Format$(dtColDate.Value, "MMM-dd-yyyy")
                If chkOldInv.Value = 1 Then
                    .TextMatrix(.Rows - 1, 2) = nsdInvoice.Text
                    .TextMatrix(.Rows - 1, 8) = nsdInvoice.BoundText
                Else
                    .TextMatrix(.Rows - 1, 2) = txtInv.Text
                End If
                .TextMatrix(.Rows - 1, 3) = txtCusAdd.Text
                .TextMatrix(.Rows - 1, 4) = cbPT.Text
                .TextMatrix(.Rows - 1, 5) = txtPayment.Text
                .TextMatrix(.Rows - 1, 6) = txtBal.Text
                .TextMatrix(.Rows - 1, 7) = txtRem.Text

                .Row = .Rows - 1
            End If
            'Increase the record count
            cCRowCount = cCRowCount + 1
        Else
            If MsgBox("Invoice payment already exist.Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
                .Row = CurrRow
                
                'Restore back the collected amount
                cCAmount = cCAmount - toNumber(Grid.TextMatrix(.RowSel, 5))
                txtTA.Text = toMoney(cCAmount)
                
                'Replace collection
                .TextMatrix(CurrRow, 1) = Format$(dtColDate.Value, "MMM-dd-yyyy")
                If chkOldInv.Value = 1 Then
                    .TextMatrix(CurrRow, 2) = nsdInvoice.Text
                    .TextMatrix(CurrRow, 8) = nsdInvoice.BoundText
                Else
                    .TextMatrix(CurrRow, 2) = txtInv.Text
                End If
                .TextMatrix(CurrRow, 3) = txtCusAdd.Text
                .TextMatrix(CurrRow, 4) = cbPT.Text
                .TextMatrix(CurrRow, 5) = txtPayment.Text
                .TextMatrix(CurrRow, 6) = txtBal.Text
                .TextMatrix(CurrRow, 7) = txtRem.Text
            Else
                Exit Sub
            End If
        End If
        'Add the amount to current load amount
        cCAmount = cCAmount + toNumber(txtPayment.Text)
        txtTA.Text = toMoney(cCAmount)
        'Highlight the current row's column
        .ColSel = 8
        'Display a remove button
        Grid_Click
        'Reset the entry fields
        ResetEntry
    End With
End Sub

Private Sub btnRemove_Click()
    'Remove selected load product
    With Grid
        'Update amount to current collection amount
        cCAmount = cCAmount - toNumber(Grid.TextMatrix(.RowSel, 5))
        txtTA.Text = toMoney(cCAmount)
        'Update the record count
        cCRowCount = cCRowCount - 1

        If .Rows = 2 Then Grid.Rows = Grid.Rows + 1
        .RemoveItem (.RowSel)
    End With

    btnRemove.Visible = False
    Grid_Click

End Sub

Private Sub chkOldInv_Click()
    If chkOldInv.Value = 1 Then
        txtInv.Visible = False
        nsdInvoice.Visible = True
        
        txtCusAdd.Visible = True
        txtBal.Visible = True
        
        Labels(5).Visible = True
        Labels(6).Visible = True
        txtPayment.Enabled = False
        
        btnCollect.Enabled = False
    Else
        txtInv.Visible = True
        nsdInvoice.Visible = False
        
        txtCusAdd.Visible = False
        txtBal.Visible = False
        
        Labels(5).Visible = False
        Labels(6).Visible = False
        txtPayment.Enabled = True
        
        btnCollect.Enabled = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    'Verify the entries
    If dcSalesman.BoundText = "" Then
        MsgBox "Please select a salesman in the list.", vbExclamation
        dcSalesman.SetFocus
        Exit Sub
    End If
    
    If dcVan.BoundText = "" Then
        MsgBox "Please select a van in the list.", vbExclamation
        dcVan.SetFocus
        Exit Sub
    End If

    If cCRowCount < 1 Then
        MsgBox "Please enter a collection first before you can save this record.", vbExclamation
        nsdInvoice.SetFocus
        Exit Sub
    End If

    If MsgBox("This save the record.Do you want to proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub


    Dim RSDetails As New Recordset
    Dim iAM As Double 'Invoice Amount Paid

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM tbl_AR_PaymentHistory WHERE VCFK=" & PK, CN, adOpenStatic, adLockOptimistic

    Screen.MousePointer = vbHourglass

    Dim c As Integer

    On Error GoTo err

    CN.BeginTrans

    'Save the record
    With rs
        If State = adStateAddMode Or State = adStatePopupMode Then
            .AddNew
            ![PK] = PK
            ![DateAdded] = Now
            ![AddedByFK] = CurrUser.USER_PK
        Else
            ![DateModified] = Now
            ![LastUserFK] = CurrUser.USER_PK
        End If
        ![CollectionNo] = txtEntry(0).Text
        ![Date] = dtpDate.Value
        ![VanFK] = dcVan.BoundText
        ![SalesmanFK] = dcSalesman.BoundText
        ![Remarks] = txtEntry(8).Text
        ![LLFK] = LLFK

        .Update

    End With

    With Grid
        'Save the details of the records
        For c = 1 To cCRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then
                
                'Check the payment type
                If .TextMatrix(c, 4) = "Post Dated Check" Then
                    Screen.MousePointer = vbDefault
            
                    frmPDCManagerAE.State = adStatePopupMode
                    frmPDCManagerAE.txtEntry(6).Text = "Payment for Invoice No. " & .TextMatrix(c, 2) & "."
                    frmPDCManagerAE.txtEntry(3).Text = .TextMatrix(c, 5)
                    frmPDCManagerAE.show vbModal
            
                    Screen.MousePointer = vbHourglass
                End If
                
                RSDetails.AddNew

                RSDetails![PK] = getIndex("tbl_AR_PaymentHistory")
                RSDetails![Date] = CDate(.TextMatrix(c, 1))
                RSDetails![PaymentType] = .TextMatrix(c, 4)
                RSDetails![Amount] = toNumber(.TextMatrix(c, 5))
                RSDetails![Balance] = toNumber(.TextMatrix(c, 6))
                RSDetails![Remarks] = .TextMatrix(c, 7)
                RSDetails![VCFK] = PK
                If toNumber(.TextMatrix(c, 8)) <> 0 Then RSDetails![InvoiceFK] = toNumber(.TextMatrix(c, 8))
        
                RSDetails.Update
                If toNumber(.TextMatrix(c, 8)) <> 0 Then
                    '***************************************************
                    '1. Get the amount paid
                    '2. Add the Amount Paid with the current pay
                    '3. Update it
                    '4. Change the status
                    '***************************************************
                    'Paid invoice
                    iAM = toNumber(getValueAt("SELECT PK,AmountPaid FROM tbl_AR_Invoice WHERE PK=" & toNumber(.TextMatrix(c, 8)), "AmountPaid"))
                    iAM = iAM + toNumber(.TextMatrix(c, 5))
                    ChangeValue CN, "tbl_AR_Invoice", "AmountPaid", iAM, True, "WHERE PK=" & toNumber(.TextMatrix(c, 8))
                    
                    If toNumber(.TextMatrix(c, 6)) <= 0 Then
                        ChangeValue CN, "tbl_AR_Invoice", "Paid", "Y", False, "WHERE PK=" & toNumber(.TextMatrix(c, 8))
                    End If
                End If

            End If

        Next c
    End With

    'Clear variables
    c = 0
    iAM = 0
    
    Set RSDetails = Nothing

    CN.CommitTrans

    HaveAction = True
    Screen.MousePointer = vbDefault

    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        Unload Me

    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If

    Exit Sub
err:
    CN.RollbackTrans
    prompt_err err, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
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
    txtEntry(0).SetFocus
End Sub

Private Sub Form_Load()
    
    'Bind the data combo
    bind_dc "SELECT * FROM tbl_AR_Van", "VanName", dcVan, "PK", True
    bind_dc "SELECT * FROM tbl_AR_Salesman", "Name", dcSalesman, "PK", True

    InitGrid

    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        frmVanCollectionAEPickFrom.show vbModal
        
        'Initialize controls
        cbPT.ListIndex = 0
        InitNSD

        'Set the recordset
         rs.Open "SELECT * FROM tbl_AR_VanCollection WHERE PK=" & PK, CN, adOpenStatic, adLockOptimistic
         dtpDate.Value = Date
         dtColDate.Value = Date
         Caption = "Create New Entry"
         cmdUsrHistory.Enabled = False
         GeneratePK
    Else
        Screen.MousePointer = vbHourglass
        'Set the recordset
        rs.Open "SELECT * FROM qry_AR_VanCollection WHERE PK=" & PK, CN, adOpenStatic, adLockOptimistic
        
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
    PK = getIndex("tbl_AR_VanCollection")
    txtEntry(0).Text = "COL" & GenerateID(PK, Format$(Date, "yyyy") & Format$(Date, "mm") & Format$(Date, "dd") & "-", "0")
End Sub

'Procedure used to initialize the grid
Private Sub InitGrid()
    cCRowCount = 0
    With Grid
        .Clear
        .ClearStructure
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .Cols = 9
        .ColSel = 8
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 2000
        .ColWidth(2) = 2000
        .ColWidth(3) = 3000
        .ColWidth(4) = 2000
        .ColWidth(5) = 1500
        .ColWidth(6) = 1500
        .ColWidth(7) = 4000
        .ColWidth(8) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Date"
        .TextMatrix(0, 2) = "Invoice No"
        .TextMatrix(0, 3) = "Customer Name"
        .TextMatrix(0, 4) = "Payment Type"
        .TextMatrix(0, 5) = "Payment"
        .TextMatrix(0, 6) = "Balance"
        .TextMatrix(0, 7) = "Remarks"
        .TextMatrix(0, 8) = "InvoiceFK"
        'Set the column alignment
        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
        .ColAlignment(2) = vbLeftJustify
        .ColAlignment(3) = vbLeftJustify
        .ColAlignment(4) = vbLeftJustify
        .ColAlignment(5) = vbLeftJustify
        .ColAlignment(6) = vbLeftJustify
        .ColAlignment(7) = vbLeftJustify
        .ColAlignment(8) = vbLeftJustify
    End With
End Sub

Private Sub ResetEntry()
    dtColDate.Value = Date
    nsdInvoice.ResetValue
    txtCusAdd.Text = ""
    txtBal.Text = "0.00"
    
    txtPayment.Text = "0.00"
    cbPT.ListIndex = 0
    txtRem.Text = ""
    
    txtInv.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        frmVanCollection.RefreshRecords
        MAIN.UpdateInfoMsg
    End If
    
    Set frmVanCollectionAE = Nothing
End Sub

Private Sub Grid_Click()
    If State = adStateEditMode Then Exit Sub
    If chkOldInv.Value = 1 Then
        If Grid.Rows = 2 And Grid.TextMatrix(1, 8) = "" Then
            btnRemove.Visible = False
        Else
            btnRemove.Visible = True
            btnRemove.Top = (Grid.CellTop + Grid.Top) - 20
            btnRemove.Left = Grid.Left + 50
        End If
    Else
        If Grid.Rows = 2 And Grid.TextMatrix(1, 5) = "" Then
            btnRemove.Visible = False
        Else
            btnRemove.Visible = True
            btnRemove.Top = (Grid.CellTop + Grid.Top) - 20
            btnRemove.Left = Grid.Left + 50
        End If
    End If

End Sub


Private Sub Grid_Scroll()
    btnRemove.Visible = False
End Sub

Private Sub Grid_SelChange()
    Grid_Click
End Sub

Private Sub nsdInvoice_Change()
    txtPayment.Text = "0.00"
    cbPT.ListIndex = 0
    txtRem.Text = ""
    
    If nsdInvoice.Text = "" Then
        txtPayment.Enabled = False
    Else
        txtPayment.Enabled = True
    End If
    
    txtCusAdd.Text = nsdInvoice.getSelValueAt(3)
    txtBal.Text = toMoney(toNumber(nsdInvoice.getSelValueAt(9)))
    txtBal.Tag = toMoney(toNumber(nsdInvoice.getSelValueAt(9)))
End Sub

Private Sub txtDate_GotFocus()
    HLText txtDate
End Sub

Private Sub txtPayment_Change()
    If chkOldInv.Value = 1 Then
        If toNumber(txtPayment.Text) > 0 Then
            btnCollect.Enabled = True
        Else
            btnCollect.Enabled = False
        End If
        
        If toNumber(txtPayment.Text) > toNumber(txtBal.Tag) Then
            txtBal.Text = "0.00"
            txtPayment.Text = toMoney(toNumber(txtBal.Tag))
            txtPayment.SelStart = Len(txtPayment.Text)
        Else
            txtBal.Text = toMoney(toNumber(txtBal.Tag) - toNumber(txtPayment.Text))
        End If
    End If
End Sub

Private Sub txtPayment_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtPayment_Validate(Cancel As Boolean)
    txtPayment.Text = toMoney(toNumber(txtPayment.Text))
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    HLText txtEntry(Index)
    If Index = 8 Then
        cmdSave.Default = False
    End If
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index > 1 And Index < 8 Then
        KeyAscii = isNumber(KeyAscii)
    End If
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
    If Index = 8 Then
        cmdSave.Default = True
    End If
End Sub

Private Sub txtEntry_Validate(Index As Integer, Cancel As Boolean)
    If Index > 1 And Index < 8 Then
        txtEntry(Index).Text = toNumber(txtEntry(Index).Text)
    End If
End Sub

'Procedure used to reset fields
Private Sub ResetFields()
    InitGrid
    ResetEntry

    dtpDate.Value = Date

    txtEntry(8).Text = ""

    txtTA.Text = "0.00"

    cCAmount = 0

    txtEntry(0).SetFocus
End Sub

'Used to display record
Private Sub DisplayForViewing()
    On Error GoTo err
    txtEntry(0).Text = rs![CollectionNo]
    txtDate.Text = Format$(rs![Date], "MMM-dd-yyyy")
    dcVan.BoundText = rs![VanFK]
    dcSalesman.BoundText = rs![SalesmanFK]
    txtEntry(8).Text = rs![Remarks]
    txtTA.Text = toMoney(toNumber(rs![Collection]))
    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_AR_VanCollectionDetails WHERE VCFK=" & PK & " ORDER BY PK ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            With Grid
                If .Rows = 2 And .TextMatrix(1, 8) = "" Then
                    .TextMatrix(1, 1) = RSDetails![Date]
                    .TextMatrix(1, 2) = RSDetails![InvoiceNo]
                    .TextMatrix(1, 3) = RSDetails![CustomerName]
                    .TextMatrix(1, 4) = RSDetails![PaymentType]
                    .TextMatrix(1, 5) = toMoney(RSDetails![Amount])
                    .TextMatrix(1, 6) = toMoney(RSDetails![Balance])
                    .TextMatrix(1, 7) = RSDetails![Remarks]
                    .TextMatrix(1, 8) = RSDetails![InvoiceFK]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![Date]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![InvoiceNo]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![CustomerName]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![PaymentType]
                    .TextMatrix(.Rows - 1, 5) = toMoney(RSDetails![Amount])
                    .TextMatrix(.Rows - 1, 6) = toMoney(RSDetails![Balance])
                    .TextMatrix(.Rows - 1, 7) = RSDetails![Remarks]
                    .TextMatrix(.Rows - 1, 8) = RSDetails![InvoiceFK]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 8
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 2
        End If
    End If

    RSDetails.Close
    'Clear variables
    Set RSDetails = Nothing

    'Disable commands
    LockInput Me, True

    picCusInfo.Visible = False
    dtpDate.Visible = False
    txtDate.Visible = True
    cmdSave.Visible = False
    btnCollect.Visible = False

    'Resize and reposition the controls
    Shape3.Top = 900
    Label11.Top = 900
    Grid.Top = 1200
    Grid.Height = 3690
    
    ctrlLiner2.Visible = False
    ctrlLiner3.Visible = False
    
    Label3.Top = 5025
    txtTA.Top = 5025
    
    Labels(4).Top = 5025
    txtEntry(8).Top = 5250
    
    ctrlLiner1.Top = 6400
    
    cmdUsrHistory.Top = 6550
    cmdCancel.Top = 6550
       
    Me.Height = 7500
    Me.Top = (Screen.Height - Me.Height) / 2

    Exit Sub
err:
    'Error if encounter a null value
    If err.Number = 94 Then Resume Next
End Sub

Private Sub InitNSD()
    'For Invoice
    With nsdInvoice
        .ClearColumn
        .AddColumn "Invoice No", 1794.89
        .AddColumn "Date", 1994.89
        .AddColumn "Sold To", 2264.88
        .AddColumn "Address", 2670.23
        .AddColumn "Discount", 1400
        .AddColumn "Total Amount", 1400
        .AddColumn "Down Payment", 1400
        .AddColumn "Amount Paid", 1400
        .AddColumn "Balance", 1400
        
        .Connection = CN.ConnectionString
        .sqlFields = "InvoiceNo,Date,SoldTo,Address,Discount,TotalAmount,DownPayment,AmountPaid,Balance,Paid,PK"
        .sqlTables = "qry_AR_Invoice"
        .sqlwCondition = "Paid='N'"
        .sqlSortOrder = "PK DESC"
        
        .BoundField = "PK"
        .PageBy = 25
        .DisplayCol = 1
        
        .setDropWindowSize 8000, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Unpaid Invoices"
        
    End With

End Sub

Private Sub txtCusAdd_GotFocus()
    HLText txtCusAdd
End Sub

Private Sub txtTA_GotFocus()
    HLText txtTA
End Sub

Private Sub txtPayment_GotFocus()
    HLText txtPayment
End Sub


