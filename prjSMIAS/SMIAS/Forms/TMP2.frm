VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form TMP2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TMP2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -600
      TabIndex        =   50
      Top             =   7200
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   53
   End
   Begin VB.PictureBox picProd 
      BorderStyle     =   0  'None
      Height          =   3090
      Left            =   -150
      ScaleHeight     =   3090
      ScaleWidth      =   11190
      TabIndex        =   4
      Top             =   975
      Width           =   11190
      Begin VB.TextBox txtEntry 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   6150
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   75
         Width           =   4890
      End
      Begin VB.TextBox txtUC 
         Height          =   285
         Left            =   1575
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   375
         Width           =   1500
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   3
         Left            =   1575
         TabIndex        =   8
         Text            =   "0"
         Top             =   1425
         Width           =   1515
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   2
         Left            =   1575
         TabIndex        =   7
         Text            =   "0"
         Top             =   1050
         Width           =   1515
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   4
         Left            =   1575
         TabIndex        =   9
         Text            =   "0"
         Top             =   1800
         Width           =   1515
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   5
         Left            =   7200
         TabIndex        =   11
         Text            =   "0"
         Top             =   1050
         Width           =   1515
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   6
         Left            =   7200
         TabIndex        =   12
         Text            =   "0"
         Top             =   1425
         Width           =   1515
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   7
         Left            =   7200
         TabIndex        =   13
         Text            =   "0"
         Top             =   1800
         Width           =   1515
      End
      Begin VB.TextBox txtLQty 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   26
         Text            =   "0"
         Top             =   2175
         Width           =   1500
      End
      Begin VB.TextBox txtVIQty 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7200
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0"
         Top             =   2175
         Width           =   1500
      End
      Begin VB.CommandButton btnPick 
         Caption         =   "Pick"
         Height          =   315
         Left            =   10200
         TabIndex        =   10
         Top             =   1050
         Width           =   840
      End
      Begin VB.CommandButton btnLoad 
         Caption         =   "Load Selected Product"
         Enabled         =   0   'False
         Height          =   315
         Left            =   9075
         TabIndex        =   18
         Top             =   2700
         Width           =   2010
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0"
         Top             =   2700
         Width           =   1500
      End
      Begin VB.TextBox txtAmount 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3975
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   2700
         Width           =   1500
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Description"
         ForeColor       =   &H0000011D&
         Height          =   240
         Index           =   2
         Left            =   4800
         TabIndex        =   49
         Top             =   75
         Width           =   1215
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Product Code"
         Height          =   240
         Index           =   3
         Left            =   300
         TabIndex        =   48
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Unit Cost(Each)"
         ForeColor       =   &H0000011D&
         Height          =   240
         Left            =   0
         TabIndex        =   47
         Top             =   375
         Width           =   1515
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Loaded Quantity"
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
         TabIndex        =   46
         Top             =   750
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Van Inventory"
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
         Left            =   5400
         TabIndex        =   45
         Top             =   750
         Width           =   2340
      End
      Begin VB.Label Label3 
         Caption         =   "(Not Available)"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   3150
         TabIndex        =   44
         Top             =   1050
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label Label4 
         Caption         =   "(Not Available)"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   3150
         TabIndex        =   43
         Top             =   1425
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Pieces"
         Height          =   240
         Index           =   14
         Left            =   300
         TabIndex        =   42
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Boxes"
         Height          =   240
         Index           =   13
         Left            =   300
         TabIndex        =   41
         Top             =   1425
         Width           =   1215
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Cases"
         Height          =   240
         Index           =   12
         Left            =   300
         TabIndex        =   40
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "(Not Available)"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   8775
         TabIndex        =   39
         Top             =   1050
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label Label7 
         Caption         =   "(Not Available)"
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   8775
         TabIndex        =   38
         Top             =   1425
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Pieces"
         Height          =   240
         Index           =   5
         Left            =   5925
         TabIndex        =   37
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Boxes"
         Height          =   240
         Index           =   6
         Left            =   5925
         TabIndex        =   36
         Top             =   1425
         Width           =   1215
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Cases"
         Height          =   240
         Index           =   7
         Left            =   5925
         TabIndex        =   35
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Loaded Qty"
         ForeColor       =   &H0000011D&
         Height          =   240
         Left            =   0
         TabIndex        =   34
         Top             =   2175
         Width           =   1515
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Van Inventory Qty"
         ForeColor       =   &H0000011D&
         Height          =   240
         Left            =   5325
         TabIndex        =   33
         Top             =   2175
         Width           =   1815
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Qty"
         ForeColor       =   &H0000011D&
         Height          =   240
         Left            =   0
         TabIndex        =   32
         Top             =   2700
         Width           =   1515
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount"
         ForeColor       =   &H0000011D&
         Height          =   240
         Left            =   3075
         TabIndex        =   31
         Top             =   2700
         Width           =   840
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000010&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   240
         Left            =   300
         Top             =   750
         Width           =   3765
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000010&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000010&
         Height          =   240
         Left            =   5250
         Top             =   750
         Width           =   5790
      End
   End
   Begin VB.CommandButton btnProdAvailable 
      Caption         =   "View Product Stocks"
      Height          =   315
      Left            =   2100
      TabIndex        =   22
      Top             =   7350
      Width           =   1830
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   225
      Picture         =   "TMP2.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Remove"
      Top             =   4725
      Visible         =   0   'False
      Width           =   275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   8100
      TabIndex        =   23
      Top             =   7350
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   9540
      TabIndex        =   24
      Top             =   7350
      Width           =   1335
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   225
      TabIndex        =   21
      Top             =   7350
      Width           =   1755
   End
   Begin VB.TextBox txtCLAmount 
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
      Left            =   9375
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0.00"
      Top             =   6825
      Width           =   1500
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1425
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   2490
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2190
      Left            =   150
      TabIndex        =   19
      Top             =   4575
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
   Begin MSDataListLib.DataCombo dcVan 
      Height          =   315
      Left            =   8400
      TabIndex        =   3
      Top             =   150
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
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
      Format          =   24510467
      CurrentDate     =   38207
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1425
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   525
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   150
      X2              =   10875
      Y1              =   4125
      Y2              =   4125
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   150
      X2              =   10875
      Y1              =   4125
      Y2              =   4125
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Van Inventory"
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
      TabIndex        =   30
      Top             =   4275
      Width           =   4365
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   " Load Amount"
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
      Left            =   7275
      TabIndex        =   29
      Top             =   6825
      Width           =   2040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   150
      X2              =   10875
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   0
      X1              =   150
      X2              =   10875
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Van"
      Height          =   240
      Index           =   4
      Left            =   7125
      TabIndex        =   28
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Inventory Date"
      Height          =   240
      Index           =   1
      Left            =   150
      TabIndex        =   27
      Top             =   525
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Inventory No"
      Height          =   240
      Index           =   0
      Left            =   75
      TabIndex        =   14
      Top             =   150
      Width           =   1290
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   150
      Top             =   4275
      Width           =   10740
   End
End
Attribute VB_Name = "TMP2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public LLFK                 As Long 'Last loading FK
Public CloseMe              As Boolean

Dim PCase                   As Long 'Pieces per case
Dim PBox                    As Long 'Pieces per box

Dim clAmount                As Currency 'Current Loading Amount
Dim clRowCount              As Integer

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset 'Main recordset for loading

Private Sub btnLoad_Click()
   
    Dim CurrRow As Integer
    
    CurrRow = getFlexPos(Grid, 11, dcProd.BoundText)
    
    'Add to grid
    With Grid
        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 11) = "" Then
                .TextMatrix(1, 1) = dcProd.Text
                .TextMatrix(1, 2) = txtEntry(1).Text
                .TextMatrix(1, 3) = txtUC.Text
                .TextMatrix(1, 4) = txtEntry(2).Text
                .TextMatrix(1, 5) = txtEntry(3).Text
                .TextMatrix(1, 6) = txtEntry(4).Text
                .TextMatrix(1, 7) = txtLQty.Text
                .TextMatrix(1, 8) = txtVIQty.Text
                .TextMatrix(1, 9) = txtQty.Text
                .TextMatrix(1, 10) = txtAmount.Text
                .TextMatrix(1, 11) = dcProd.BoundText
                
            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = dcProd.Text
                .TextMatrix(.Rows - 1, 2) = txtEntry(1).Text
                .TextMatrix(.Rows - 1, 3) = txtUC.Text
                .TextMatrix(.Rows - 1, 4) = txtEntry(2).Text
                .TextMatrix(.Rows - 1, 5) = txtEntry(3).Text
                .TextMatrix(.Rows - 1, 6) = txtEntry(4).Text
                .TextMatrix(.Rows - 1, 7) = txtLQty.Text
                .TextMatrix(.Rows - 1, 8) = txtVIQty.Text
                .TextMatrix(.Rows - 1, 9) = txtQty.Text
                .TextMatrix(.Rows - 1, 10) = txtAmount.Text
                .TextMatrix(.Rows - 1, 11) = dcProd.BoundText
                
                .Row = .Rows - 1
            End If
            'Increase the record count
            clRowCount = clRowCount + 1
        Else
            'Perform if the record already exist
            .Row = CurrRow
            
            .TextMatrix(CurrRow, 1) = dcProd.Text
            .TextMatrix(CurrRow, 2) = txtEntry(1).Text
            .TextMatrix(CurrRow, 3) = txtUC.Text
            .TextMatrix(CurrRow, 4) = txtEntry(2).Text + toNumber(.TextMatrix(CurrRow, 4))
            .TextMatrix(CurrRow, 5) = txtEntry(3).Text + toNumber(.TextMatrix(CurrRow, 5))
            .TextMatrix(CurrRow, 6) = txtEntry(4).Text + toNumber(.TextMatrix(CurrRow, 6))
            .TextMatrix(CurrRow, 7) = toNumber(txtLQty.Text) + toNumber(.TextMatrix(CurrRow, 7))
            .TextMatrix(CurrRow, 8) = toNumber(txtVIQty.Text) + toNumber(.TextMatrix(CurrRow, 8))
            .TextMatrix(CurrRow, 9) = toNumber(txtQty.Text) + toNumber(.TextMatrix(CurrRow, 9))
            .TextMatrix(CurrRow, 10) = toNumber(txtAmount.Text) + toNumber(.TextMatrix(CurrRow, 10))
            .TextMatrix(CurrRow, 11) = dcProd.BoundText
            
        End If
        'Add the amount to current load amount
        clAmount = clAmount + toNumber(txtAmount.Text)
        txtCLAmount.Text = Format$(clAmount, "#,##0.00")
        'Highlight the current row's column
        .ColSel = 11
        'Display a remove button
        Grid_Click
        'Reset the entry fields
        ResetEntry
    End With
End Sub

Private Sub btnProdAvailable_Click()
    'Display Product Stock Info
    frmStockViewer.show vbModal
End Sub

Private Sub btnRemove_Click()
    'Remove selected load product
    With Grid
        'Update amount to current load amount
        clAmount = clAmount - toNumber(Grid.TextMatrix(.RowSel, 10))
        txtCLAmount.Text = Format$(clAmount, "#,##0.00")
        'Update the record count
        clRowCount = clRowCount - 1
        
        If .Rows = 2 Then Grid.Rows = Grid.Rows + 1
        .RemoveItem (.RowSel)
    End With

    btnRemove.Visible = False
    Grid_Click
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    'Verify the entries
    If dcVan.BoundText = "" Then
        MsgBox "Please select a van in the list.", vbExclamation
        dcVan.SetFocus
        Exit Sub
    End If
    If clRowCount < 1 Then
        MsgBox "Please load a product first before you can save this record.", vbExclamation
        dcProd.SetFocus
        Exit Sub
    End If
    
    If MsgBox("This save the record.Do you want to proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    Dim RSDetails As New Recordset
    Dim EntryIsOK As Boolean
    Dim ProdPK As Long 'Product Primary Key
    Dim tC As Long 'Temporary Case - Based on actual product quantity
    Dim tB As Long 'Temporary Box --^
    Dim tP As Long 'Temporary Pieces --^
    
    EntryIsOK = True
    
    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM tbl_IC_LoadingDetails WHERE LoadingFK=" & PK, CN, adOpenStatic, adLockOptimistic
    
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
        ![LoadingNo] = txtEntry(0).Text
        ![Date] = dtpDate.Value
        ![VanFK] = dcVan.BoundText
        
        .Update
    End With
    
    With Grid
        'Save the details of the records
        For c = 1 To clRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then
            
                ProdPK = toNumber(.TextMatrix(c, 11))
                
                tC = toNumber(getValueAt("SELECT PK,Cases FROM tbl_IC_Products WHERE PK=" & ProdPK, "Cases"))
                tB = toNumber(getValueAt("SELECT PK,Boxes FROM tbl_IC_Products WHERE PK=" & ProdPK, "Boxes"))
                tP = toNumber(getValueAt("SELECT PK,Pieces FROM tbl_IC_Products WHERE PK=" & ProdPK, "Pieces"))
                
                If toNumber(.TextMatrix(c, 4)) > tC Then EntryIsOK = False: .Col = 4: .CellForeColor = &HFF&: .CellFontBold = True
                If toNumber(.TextMatrix(c, 5)) > tB Then EntryIsOK = False: .Col = 5: .CellForeColor = &HFF&: .CellFontBold = True
                If toNumber(.TextMatrix(c, 6)) > tP Then EntryIsOK = False: .Col = 6: .CellForeColor = &HFF&: .CellFontBold = True
                
                RSDetails.AddNew
                
                RSDetails![PK] = getIndex("tbl_IC_LoadingDetails")
                
                RSDetails![LoadingFK] = PK
                RSDetails![ProductFK] = ProdPK
                RSDetails![UnitCost(Each)] = .TextMatrix(c, 3)
                RSDetails![Cases] = .TextMatrix(c, 4)
                RSDetails![Boxes] = .TextMatrix(c, 5)
                RSDetails![Pieces] = .TextMatrix(c, 6)
                RSDetails![QtyLoad] = .TextMatrix(c, 7)
                RSDetails![VanInv] = .TextMatrix(c, 8)
                
                RSDetails.Update
                
                'Update stock value
                ChangeValue CN, "tbl_IC_Products", "Cases", tC - toNumber(.TextMatrix(c, 4)), True, "WHERE PK=" & ProdPK
                ChangeValue CN, "tbl_IC_Products", "Boxes", tB - toNumber(.TextMatrix(c, 5)), True, "WHERE PK=" & ProdPK
                ChangeValue CN, "tbl_IC_Products", "Pieces", tP - toNumber(.TextMatrix(c, 6)), True, "WHERE PK=" & ProdPK

            End If

        Next c
    End With
    
    'Clear variables
    c = 0
    ProdPK = 0
    tC = 0
    tB = 0
    tP = 0
    Set RSDetails = Nothing
    
    If EntryIsOK = True Then
        CN.CommitTrans
    Else
        CN.RollbackTrans
        MsgBox "Some product/s have not enough quantity to serve for this loading." & vbCrLf & _
               "Please check the stock value of the loaded products with red color in the list.", vbExclamation
        Grid.Row = 1
        Grid.Col = 0
        'Grid.ColSel = 11
        Grid.SetFocus
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    HaveAction = True
    Screen.MousePointer = vbDefault

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

Private Sub dcProd_Click(Area As Integer)
    On Error Resume Next
    If Area = 2 Then
        If dcProd.BoundText <> "" Then
            ResetEntry
            DiplayProdInfo
        End If
    End If
End Sub

Private Sub Form_Activate()
    If CloseMe = True Then Unload Me
End Sub

Private Sub Form_Load()
    dtpDate.Value = Date
    
    'Bind the data combo
    bind_dc "SELECT * FROM tbl_AR_Van", "VanName", dcVan, "PK", True
    
    InitGrid
    
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        frmVanInventoryAEPickFrom.show vbModal
        
        'Set the recordset
        rs.Open "SELECT * FROM tbl_IC_Loading WHERE PK=" & PK, CN, adOpenStatic, adLockOptimistic
        'Bind the combo
        bind_dc "SELECT * FROM tbl_IC_Products", "ProductCode", dcProd, "PK", True
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        GeneratePK
        DiplayProdInfo
    Else
        Screen.MousePointer = vbHourglass
        'Set the recordset
        rs.Open "SELECT * FROM qry_IC_Loading WHERE PK=" & PK, CN, adOpenStatic, adLockOptimistic

        Caption = "View Record"
        cmdCancel.Caption = "Close"
        cmdUsrHistory.Enabled = True
        btnProdAvailable.Enabled = False
        txtEntry(0).Width = txtDate.Width
        DisplayForViewing
        Me.show
        MsgBox "This is use for viewing the record only." & vbCrLf & _
               "You cannot perform any changes in this form." & vbCrLf & vbCrLf & _
               "Note:If you have mistake in adding this record then " & vbCrLf & _
               "void this record and re-enter.", vbExclamation
        txtEntry(0).SetFocus
        Screen.MousePointer = vbDefault
    End If
    
End Sub

'Procedure used to generate PK
Private Sub GeneratePK()
    PK = getIndex("tbl_IC_Loading")
    txtEntry(0).Text = GenerateID(PK, Format$(Date, "yyyy") & Format$(Date, "mm") & Format$(Date, "dd") & "-", "0")
End Sub

'Procedure used to initialize the grid
Private Sub InitGrid()
    clRowCount = 0
    With Grid
        .Clear
        .ClearStructure
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .Cols = 13
        .ColSel = 11
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 2025
        .ColWidth(2) = 2505
        .ColWidth(3) = 1300
        .ColWidth(4) = 1300
        .ColWidth(5) = 1300
        .ColWidth(6) = 1300
        .ColWidth(7) = 1300
        .ColWidth(8) = 1300
        .ColWidth(9) = 1300
        .ColWidth(10) = 1300
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Product Code"
        .TextMatrix(0, 2) = "Description"
        .TextMatrix(0, 3) = "Unit Cost(Each)"
        .TextMatrix(0, 4) = "Cases"
        .TextMatrix(0, 5) = "Boxes"
        .TextMatrix(0, 6) = "Pieces"
        .TextMatrix(0, 7) = "Qty Load"
        .TextMatrix(0, 8) = "Van Inv"
        .TextMatrix(0, 9) = "Total Load"
        .TextMatrix(0, 10) = "Amount"
        .TextMatrix(0, 11) = "ProductFK"
        .TextMatrix(0, 12) = "PK"
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
        .ColAlignment(9) = vbLeftJustify
        .ColAlignment(10) = vbLeftJustify
        .ColAlignment(11) = vbLeftJustify
        .ColAlignment(12) = vbLeftJustify
    End With
End Sub

'Procedure used to display product information
Private Sub DiplayProdInfo()
    Screen.MousePointer = vbHourglass
    
    Dim rsPI As New Recordset
    
    With rsPI
        .CursorLocation = adUseClient
        
        .Open "SELECT * FROM tbl_IC_Products WHERE PK =" & dcProd.BoundText, CN, adOpenStatic, adLockReadOnly
        
        txtEntry(1).Text = ![Description]
        txtUC.Text = toMoney(toNumber(![UnitCost]))
        PCase = ![PiecesPerCase]
        PBox = ![PiecesPerBox]
        
    End With
    
    Set rsPI = Nothing
    
    If PCase = 0 Then
        Label3.Visible = True
        Label5.Visible = True
        'Loaded Entry
        txtEntry(2).BackColor = &HE6FFFF
        txtEntry(2).ForeColor = &H0&
        txtEntry(2).Locked = True
        'Van Inventory Entry
        txtEntry(5).BackColor = &HE6FFFF
        txtEntry(5).ForeColor = &H0&
        txtEntry(5).Locked = True
    Else
        Label3.Visible = False
        Label5.Visible = False
        'Loaded Entry
        txtEntry(2).BackColor = &H80000005
        txtEntry(2).ForeColor = &H80000008
        txtEntry(2).Locked = False
        'Van Inventory Entry
        txtEntry(5).BackColor = &H80000005
        txtEntry(5).ForeColor = &H80000008
        txtEntry(5).Locked = False
    End If
    
    If PBox = 0 Then
        Label4.Visible = True
        Label7.Visible = True
        'Loaded Entry
        txtEntry(3).BackColor = &HE6FFFF
        txtEntry(3).ForeColor = &H0&
        txtEntry(3).Locked = True
        'Van Inventory Entry
        txtEntry(6).BackColor = &HE6FFFF
        txtEntry(6).ForeColor = &H0&
        txtEntry(6).Locked = True
    Else
        Label4.Visible = False
        Label7.Visible = False
        'Loaded Entry
        txtEntry(3).BackColor = &H80000005
        txtEntry(3).ForeColor = &H80000008
        txtEntry(3).Locked = False
        'Van Inventory Entry
        txtEntry(6).BackColor = &H80000005
        txtEntry(6).ForeColor = &H80000008
        txtEntry(6).Locked = False
    End If
    
    Screen.MousePointer = vbDefault
End Sub


Private Sub ResetEntry()
    txtEntry(2).Text = "0"
    txtEntry(3).Text = "0"
    txtEntry(4).Text = "0"
    
    txtEntry(5).Text = "0"
    txtEntry(6).Text = "0"
    txtEntry(7).Text = "0"
        
    txtLQty.Text = "0"
    txtVIQty.Text = "0"
   
    txtQty.Text = "0"
    txtAmount.Text = "0.00"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        frmVanInventory.RefreshRecords
        MAIN.UpdateInfoMsg
    End If
    
    'Clean up all used variables
    Set rs = Nothing
    
    PK = 0

    PCase = 0
    PBox = 0

    clAmount = 0
    clRowCount = 0

    HaveAction = False
    
    Set frmVanInventoryAE = Nothing
End Sub

Private Sub Grid_Click()
    If State = adStateEditMode Then Exit Sub
    If Grid.Rows = 2 And Grid.TextMatrix(1, 11) = "" Then
        btnRemove.Visible = False
    Else
        btnRemove.Visible = True
        btnRemove.Top = (Grid.CellTop + Grid.Top) - 20
        btnRemove.Left = Grid.Left + 50
    End If
End Sub


Private Sub Grid_Scroll()
    btnRemove.Visible = False
End Sub

Private Sub Grid_SelChange()
    Grid_Click
End Sub

Private Sub txtAmount_GotFocus()
    HLText txtAmount
End Sub

Private Sub txtCLAmount_GotFocus()
    HLText txtCLAmount
End Sub

Private Sub txtDate_GotFocus()
    HLText txtDate
End Sub

Private Sub txtEntry_Change(Index As Integer)
    If Index > 1 And Index < 8 Then
        txtLQty.Text = (toNumber(txtEntry(2).Text) * PCase) + (toNumber(txtEntry(3).Text) * PBox) + toNumber(txtEntry(4).Text)
        txtVIQty.Text = (toNumber(txtEntry(5).Text) * PCase) + (toNumber(txtEntry(6).Text) * PBox) + toNumber(txtEntry(7).Text)
        txtQty.Text = toNumber(txtLQty.Text) + toNumber(txtVIQty.Text)
        
        txtAmount.Text = toMoney(toNumber(txtQty.Text) * toNumber(txtUC.Text))
    End If
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index > 1 And Index < 8 Then
        KeyAscii = isNumber(KeyAscii)
    End If
End Sub

Private Sub txtEntry_Validate(Index As Integer, Cancel As Boolean)
    If Index > 1 And Index < 8 Then
        txtEntry(Index).Text = toNumber(txtEntry(Index).Text)
    End If
End Sub

Private Sub txtLQty_GotFocus()
    HLText txtLQty
End Sub

Private Sub txtQty_Change()
    If toNumber(txtQty.Text) < 1 Then
        btnLoad.Enabled = False
    Else
        btnLoad.Enabled = True
    End If
End Sub

Private Sub txtQty_GotFocus()
    HLText txtQty
End Sub

Private Sub txtUC_Change()
    txtLQty.Text = (toNumber(txtEntry(2).Text) * PCase) + (toNumber(txtEntry(3).Text) * PBox) + toNumber(txtEntry(4).Text)
    txtVIQty.Text = (toNumber(txtEntry(5).Text) * PCase) + (toNumber(txtEntry(6).Text) * PBox) + toNumber(txtEntry(7).Text)
    txtQty.Text = toNumber(txtLQty.Text) + toNumber(txtVIQty.Text)
    
    txtAmount.Text = toMoney(toNumber(txtQty.Text) * toNumber(txtUC.Text))
End Sub

Private Sub txtUC_GotFocus()
    HLText txtUC
End Sub

Private Sub txtUC_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtUC_Validate(Cancel As Boolean)
    txtUC.Text = toMoney(toNumber(txtUC.Text))
End Sub

Private Sub txtVIQty_GotFocus()
    HLText txtVIQty
End Sub

'Procedure used to reset fields
Private Sub ResetFields()
    clearText Me
    InitGrid
    
    dtpDate.Value = Date
    
    ResetEntry
    
    clAmount = 0
    
    txtUC.Text = "0.00"
    txtCLAmount.Text = "0.00"
    
    dcVan.BoundText = RightSplitUF(dcVan.Tag)
    dcProd.BoundText = RightSplitUF(dcProd.Tag)
    DiplayProdInfo
    
    dtpDate.SetFocus
End Sub

'Used to display record
Private Sub DisplayForViewing()
    txtEntry(0).Text = rs![LoadingNo]
    txtDate.Text = Format$(rs![Date], "MMM-dd-yyyy")
    dcVan.BoundText = rs![VanFK]
    txtCLAmount.Text = toMoney(rs![TotalAmount])
    'Display the details
    Dim RSDetails As New Recordset
    
    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_IC_LoadingDetails WHERE LoadingFK=" & PK & " ORDER BY PK ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            With Grid
                If .Rows = 2 And .TextMatrix(1, 11) = "" Then
                    .TextMatrix(1, 1) = RSDetails![ProductCode]
                    .TextMatrix(1, 2) = RSDetails![Description]
                    .TextMatrix(1, 3) = toMoney(RSDetails![UnitCost(Each)])
                    .TextMatrix(1, 4) = RSDetails![Cases]
                    .TextMatrix(1, 5) = RSDetails![Boxes]
                    .TextMatrix(1, 6) = RSDetails![Pieces]
                    .TextMatrix(1, 7) = RSDetails![QtyLoad]
                    .TextMatrix(1, 8) = RSDetails![VanInv]
                    .TextMatrix(1, 9) = RSDetails![TotalLoad]
                    .TextMatrix(1, 10) = toMoney(RSDetails![Amount])
                    .TextMatrix(1, 11) = RSDetails![ProductFK]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![ProductCode]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Description]
                    .TextMatrix(.Rows - 1, 3) = toMoney(RSDetails![UnitCost(Each)])
                    .TextMatrix(.Rows - 1, 4) = RSDetails![Cases]
                    .TextMatrix(.Rows - 1, 5) = RSDetails![Boxes]
                    .TextMatrix(.Rows - 1, 6) = RSDetails![Pieces]
                    .TextMatrix(.Rows - 1, 7) = RSDetails![QtyLoad]
                    .TextMatrix(.Rows - 1, 8) = RSDetails![VanInv]
                    .TextMatrix(.Rows - 1, 9) = RSDetails![TotalLoad]
                    .TextMatrix(.Rows - 1, 10) = toMoney(RSDetails![Amount])
                    .TextMatrix(.Rows - 1, 11) = RSDetails![ProductFK]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 11
    End If
    
    RSDetails.Close
    
    'Disable commands
    LockInput Me, True
    
    dtpDate.Visible = False
    txtDate.Visible = True
    picProd.Visible = False
    cmdSave.Visible = False
    btnLoad.Visible = False
    
    'Resize and reposition the controls
    Shape3.Top = 900
    Label11.Top = 900
    Grid.Top = 1200
    Grid.Height = 3565
    txtCLAmount.Top = txtCLAmount.Top - 2000
    Label9.Top = Label9.Top - 2000
    cmdUsrHistory.Top = cmdUsrHistory.Top - 2000
    btnProdAvailable.Top = btnProdAvailable.Top - 2000
    cmdSave.Top = cmdSave.Top - 2000
    cmdCancel.Top = cmdCancel.Top - 2000
    ctrlLiner1.Top = cmdSave.Top - 150
    Me.Height = Me.Height - 2000
    Me.Top = (Screen.Height - Me.Height) / 2
    Line1(0).Visible = False
    Line1(1).Visible = False
    Line2(0).Visible = False
    Line2(1).Visible = False
    'Clear variables
    Set RSDetails = Nothing
End Sub

