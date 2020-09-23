VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmVanInventoryAE 
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
   Icon            =   "frmVanInventoryAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   11055
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtVan 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   150
      Width           =   3075
   End
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -600
      TabIndex        =   37
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
      TabIndex        =   3
      Top             =   975
      Width           =   11190
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
         Left            =   9525
         Locked          =   -1  'True
         TabIndex        =   46
         Text            =   "0"
         Top             =   1875
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
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   9
         Left            =   9525
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "0"
         Top             =   2250
         Width           =   1500
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   2
         Left            =   9525
         TabIndex        =   43
         Text            =   "0"
         Top             =   1500
         Width           =   1515
      End
      Begin VB.TextBox txtReadOnly 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   9525
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "0"
         Top             =   1125
         Width           =   1500
      End
      Begin VB.TextBox txtReadOnly 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "0.00"
         Top             =   1500
         Width           =   1500
      End
      Begin VB.TextBox txtProdDesc 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   6150
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   75
         Width           =   4890
      End
      Begin VB.TextBox txtUC 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   5
         Text            =   "0.00"
         Top             =   375
         Width           =   1500
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   0
         Left            =   1575
         TabIndex        =   6
         Text            =   "0"
         Top             =   1125
         Width           =   1515
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   1
         Left            =   1575
         TabIndex        =   7
         Text            =   "0"
         Top             =   1875
         Width           =   1515
      End
      Begin VB.TextBox txtReadOnly 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   5175
         Locked          =   -1  'True
         TabIndex        =   8
         Text            =   "0"
         Top             =   1125
         Width           =   1515
      End
      Begin VB.TextBox txtReadOnly 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   5175
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "0"
         Top             =   1500
         Width           =   1515
      End
      Begin VB.TextBox txtReadOnly 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   5175
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0"
         Top             =   1875
         Width           =   1515
      End
      Begin VB.TextBox txtReadOnly 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   1575
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   2250
         Width           =   1500
      End
      Begin VB.CommandButton btnAddInv 
         Caption         =   "Add To Inventory"
         Height          =   315
         Left            =   9225
         TabIndex        =   14
         Top             =   2700
         Width           =   1785
      End
      Begin VB.TextBox txtReadOnly 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   5175
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0"
         Top             =   2250
         Width           =   1500
      End
      Begin VB.TextBox txtReadOnly 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   5175
         Locked          =   -1  'True
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   2625
         Width           =   1500
      End
      Begin ctrlNSDataCombo.NSDataCombo nsdProduct 
         Height          =   315
         Left            =   1575
         TabIndex        =   38
         Top             =   0
         Width           =   2490
         _ExtentX        =   4392
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
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Short"
         ForeColor       =   &H0000011D&
         Height          =   240
         Left            =   7950
         TabIndex        =   47
         Top             =   1875
         Width           =   1515
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Over"
         ForeColor       =   &H0000011D&
         Height          =   240
         Left            =   7950
         TabIndex        =   45
         Top             =   2250
         Width           =   1515
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Actual Van Inventory"
         ForeColor       =   &H0000011D&
         Height          =   240
         Left            =   7575
         TabIndex        =   42
         Top             =   1500
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Unsold Qty"
         ForeColor       =   &H0000011D&
         Height          =   240
         Left            =   7950
         TabIndex        =   41
         Top             =   1125
         Width           =   1515
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Description"
         ForeColor       =   &H0000011D&
         Height          =   240
         Index           =   2
         Left            =   4800
         TabIndex        =   36
         Top             =   75
         Width           =   1215
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Product Code"
         Height          =   240
         Index           =   3
         Left            =   300
         TabIndex        =   35
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Unit Cost(Each)"
         ForeColor       =   &H0000011D&
         Height          =   240
         Left            =   0
         TabIndex        =   34
         Top             =   375
         Width           =   1515
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Return Stocks"
         Height          =   240
         Index           =   14
         Left            =   300
         TabIndex        =   33
         Top             =   1875
         Width           =   1215
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "B.O. Amount"
         Height          =   240
         Index           =   13
         Left            =   300
         TabIndex        =   32
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "B.O."
         Height          =   240
         Index           =   12
         Left            =   300
         TabIndex        =   31
         Top             =   1125
         Width           =   1215
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Sold Pieces"
         Height          =   240
         Index           =   5
         Left            =   3900
         TabIndex        =   30
         Top             =   1875
         Width           =   1215
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Sold Boxes"
         Height          =   240
         Index           =   6
         Left            =   3900
         TabIndex        =   29
         Top             =   1500
         Width           =   1215
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Sold Cases"
         Height          =   240
         Index           =   7
         Left            =   3900
         TabIndex        =   28
         Top             =   1125
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "R.S. Amount"
         ForeColor       =   &H0000011D&
         Height          =   240
         Left            =   0
         TabIndex        =   27
         Top             =   2250
         Width           =   1515
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Qty Sold"
         ForeColor       =   &H0000011D&
         Height          =   240
         Left            =   3600
         TabIndex        =   26
         Top             =   2250
         Width           =   1515
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Qty Sold Amount"
         ForeColor       =   &H0000011D&
         Height          =   240
         Left            =   3000
         TabIndex        =   25
         Top             =   2625
         Width           =   2115
      End
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   225
      Picture         =   "frmVanInventoryAE.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   20
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
      TabIndex        =   18
      Top             =   7350
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   9540
      TabIndex        =   19
      Top             =   7350
      Width           =   1335
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   225
      TabIndex        =   17
      Top             =   7350
      Width           =   1755
   End
   Begin VB.TextBox txtcInvAmount 
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
      TabIndex        =   16
      Text            =   "0.00"
      Top             =   6825
      Width           =   1500
   End
   Begin VB.TextBox txtVInvNo 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
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
      TabIndex        =   15
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
      Format          =   19595267
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
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Van"
      Height          =   240
      Index           =   4
      Left            =   6525
      TabIndex        =   49
      Top             =   150
      Width           =   1215
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
      TabIndex        =   24
      Top             =   4275
      Width           =   4365
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Sold Amount"
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
      TabIndex        =   23
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
      Caption         =   "Inventory Date"
      Height          =   240
      Index           =   1
      Left            =   150
      TabIndex        =   22
      Top             =   525
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Inventory  No"
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
      Left            =   75
      TabIndex        =   11
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
Attribute VB_Name = "frmVanInventoryAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public LLFK                 As Long 'Last loading FK
Public LLDate               As String
Public CloseMe              As Boolean


Dim cInvAmount                As Currency 'Current Loading Amount
Dim cInvRowCount              As Integer

Dim load_qty                  As Integer
Dim load_vanqty               As Integer

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset 'Main recordset for loading

Private Sub btnAddINV_Click()
   
    Dim CurrRow As Integer
    
    CurrRow = getFlexPos(Grid, 17, nsdProduct.BoundText)
    
    'Add to grid
    With Grid
        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 17) = "" Then
                .TextMatrix(1, 1) = nsdProduct.Text
                .TextMatrix(1, 2) = txtProdDesc.Text
                .TextMatrix(1, 3) = txtUC.Text
                .TextMatrix(1, 4) = txtEntry(0).Text
                .TextMatrix(1, 5) = txtReadOnly(0).Text
                .TextMatrix(1, 6) = txtEntry(1).Text
                .TextMatrix(1, 7) = txtReadOnly(1).Text
                .TextMatrix(1, 8) = txtReadOnly(2).Text
                .TextMatrix(1, 9) = txtReadOnly(3).Text
                .TextMatrix(1, 10) = txtReadOnly(4).Text
                .TextMatrix(1, 11) = txtReadOnly(5).Text
                .TextMatrix(1, 12) = txtReadOnly(6).Text
                .TextMatrix(1, 13) = txtReadOnly(7).Text
                .TextMatrix(1, 14) = txtEntry(2).Text
                .TextMatrix(1, 15) = txtReadOnly(8).Text
                .TextMatrix(1, 16) = txtReadOnly(9).Text
                .TextMatrix(1, 17) = nsdProduct.BoundText
                .TextMatrix(1, 18) = load_qty
                .TextMatrix(1, 19) = load_vanqty
            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = nsdProduct.Text
                .TextMatrix(.Rows - 1, 2) = txtProdDesc.Text
                .TextMatrix(.Rows - 1, 3) = txtUC.Text
                .TextMatrix(.Rows - 1, 4) = txtEntry(0).Text
                .TextMatrix(.Rows - 1, 5) = txtReadOnly(0).Text
                .TextMatrix(.Rows - 1, 6) = txtEntry(1).Text
                .TextMatrix(.Rows - 1, 7) = txtReadOnly(1).Text
                .TextMatrix(.Rows - 1, 8) = txtReadOnly(2).Text
                .TextMatrix(.Rows - 1, 9) = txtReadOnly(3).Text
                .TextMatrix(.Rows - 1, 10) = txtReadOnly(4).Text
                .TextMatrix(.Rows - 1, 11) = txtReadOnly(5).Text
                .TextMatrix(.Rows - 1, 12) = txtReadOnly(6).Text
                .TextMatrix(.Rows - 1, 13) = txtReadOnly(7).Text
                .TextMatrix(.Rows - 1, 14) = txtEntry(2).Text
                .TextMatrix(.Rows - 1, 15) = txtReadOnly(8).Text
                .TextMatrix(.Rows - 1, 16) = txtReadOnly(9).Text
                .TextMatrix(.Rows - 1, 17) = nsdProduct.BoundText
                .TextMatrix(.Rows - 1, 18) = load_qty
                .TextMatrix(.Rows - 1, 19) = load_vanqty
                
                .Row = .Rows - 1
            End If
            'Increase the record count
            cInvRowCount = cInvRowCount + 1
        Else
            If MsgBox("Product already exist in the inventory.Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
                'Perform if the record already exist
                .Row = CurrRow
                
                .TextMatrix(CurrRow, 1) = nsdProduct.Text
                .TextMatrix(CurrRow, 2) = txtProdDesc.Text
                .TextMatrix(CurrRow, 3) = txtUC.Text
                .TextMatrix(CurrRow, 4) = txtEntry(0).Text
                .TextMatrix(CurrRow, 5) = txtReadOnly(0).Text
                .TextMatrix(CurrRow, 6) = txtEntry(1).Text
                .TextMatrix(CurrRow, 7) = txtReadOnly(1).Text
                .TextMatrix(CurrRow, 8) = txtReadOnly(2).Text
                .TextMatrix(CurrRow, 9) = txtReadOnly(3).Text
                .TextMatrix(CurrRow, 10) = txtReadOnly(4).Text
                .TextMatrix(CurrRow, 11) = txtReadOnly(5).Text
                .TextMatrix(CurrRow, 12) = txtReadOnly(6).Text
                .TextMatrix(CurrRow, 13) = txtReadOnly(7).Text
                .TextMatrix(CurrRow, 14) = txtEntry(2).Text
                .TextMatrix(CurrRow, 15) = txtReadOnly(8).Text
                .TextMatrix(CurrRow, 16) = txtReadOnly(9).Text
                .TextMatrix(CurrRow, 17) = nsdProduct.BoundText
                .TextMatrix(CurrRow, 18) = load_qty
                .TextMatrix(CurrRow, 19) = load_vanqty
            Else
                Exit Sub
            End If
        End If
        'Add the amount to current van inventory amount
        cInvAmount = cInvAmount + toNumber(txtReadOnly(6).Text)
        txtcInvAmount.Text = Format$(cInvAmount, "#,##0.00")
        'Highlight the current row's column
        .ColSel = 17
        'Display a remove button
        Grid_Click
        'Reset the entry fields
        ResetEntry
    End With
End Sub

Private Sub btnRemove_Click()
    'Remove selected load product
    With Grid
        'Update amount to current inventory amount
        cInvAmount = cInvAmount - toNumber(Grid.TextMatrix(.RowSel, 12))
        txtcInvAmount.Text = Format$(cInvAmount, "#,##0.00")
        'Update the record count
        cInvRowCount = cInvRowCount - 1
        
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
    If cInvRowCount < 1 Then
        MsgBox "Please inventory a product first before you can save this record.", vbExclamation
        nsdProduct.SetFocus
        Exit Sub
    End If

    If MsgBox("This will save the record.Do you want to proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Screen.MousePointer = vbHourglass

    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM tbl_IC_VanInvDetails WHERE VanInvFK=" & PK, CN, adOpenStatic, adLockOptimistic

    Dim c As Integer, BO As Long, OLD_BO As Long

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
        ![VanInventoryNo] = txtVInvNo.Text
        ![Date] = dtpDate.Value
        ![VanFK] = toNumber(txtVan.Tag)
        ![LLFK] = LLFK

        .Update
    End With

    With Grid
        'Save the details of the records
        For c = 1 To cInvRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then

                RSDetails.AddNew

                RSDetails![PK] = getIndex("tbl_IC_VanInvDetails")

                RSDetails![VanInvFK] = PK
                RSDetails![ProductFK] = toNumber(.TextMatrix(c, 17))
                RSDetails![UnitCost(Each)] = toNumber(.TextMatrix(c, 3))
                RSDetails![BO] = toNumber(.TextMatrix(c, 4))
                BO = toNumber(.TextMatrix(c, 4))
                RSDetails![ReturnStock] = toNumber(.TextMatrix(c, 5))
                RSDetails![SoldPieces] = toNumber(.TextMatrix(c, 10))
                RSDetails![SoldBoxes] = toNumber(.TextMatrix(c, 9))
                RSDetails![SoldCases] = toNumber(.TextMatrix(c, 8))
                RSDetails![TotalQty] = toNumber(.TextMatrix(c, 11))
                RSDetails![UnsoldQty] = toNumber(.TextMatrix(c, 13))
                RSDetails![ActualVanInv] = toNumber(.TextMatrix(c, 14))
                RSDetails![Short] = toNumber(.TextMatrix(c, 15))
                RSDetails![Over] = toNumber(.TextMatrix(c, 16))
                RSDetails![LoadQty] = toNumber(.TextMatrix(c, 18))
                RSDetails![LoadVanInv] = toNumber(.TextMatrix(c, 19))

                RSDetails.Update
                
                If BO > 0 Then
                    'Update product BO
                    OLD_BO = toNumber(getValueAt("SELECT PK,BO FROM tbl_IC_Products WHERE PK=" & toNumber(.TextMatrix(c, 17)), "BO"))
                    ChangeValue CN, "tbl_IC_Products", "BO", OLD_BO + BO, True, "WHERE PK=" & toNumber(.TextMatrix(c, 17))
                End If

            End If

        Next c
    End With
    
    'Clear variables
    BO = 0: OLD_BO = 0
    c = 0
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
    txtVInvNo.SetFocus
End Sub

Private Sub Form_Load()
    dtpDate.Value = Date
      
    InitGrid
    
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        frmLoadingAEPickFrom.FOR_VAN_INV = True
        frmLoadingAEPickFrom.show vbModal
        
        'Set the recordset
        rs.Open "SELECT * FROM tbl_IC_VanInv WHERE PK=" & PK, CN, adOpenStatic, adLockOptimistic
        InitNSD
        
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        GeneratePK
    Else
        Screen.MousePointer = vbHourglass
        'Set the recordset
        rs.Open "SELECT * FROM qry_IC_VanInv WHERE PK=" & PK, CN, adOpenStatic, adLockOptimistic

        Caption = "View Record"
        cmdCancel.Caption = "Close"
        cmdUsrHistory.Enabled = True
        txtVInvNo.Width = txtDate.Width
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
    PK = getIndex("tbl_IC_VanInv")
    txtVInvNo.Text = GenerateID(PK, Format$(Date, "yyyy") & Format$(Date, "mm") & Format$(Date, "dd") & "-", "0")
End Sub

'Procedure used to initialize the grid
Private Sub InitGrid()
    cInvRowCount = 0
    With Grid
        .Clear
        .ClearStructure
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .Cols = 20
        .ColSel = 16
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
        .ColWidth(11) = 1300
        .ColWidth(12) = 1300
        .ColWidth(13) = 1300
        .ColWidth(14) = 1300
        .ColWidth(15) = 1300
        .ColWidth(16) = 1300
        .ColWidth(17) = 0
        .ColWidth(18) = 0
        .ColWidth(19) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Product Code"
        .TextMatrix(0, 2) = "Description"
        .TextMatrix(0, 3) = "Unit Cost(Each)"
        .TextMatrix(0, 4) = "B.O."
        .TextMatrix(0, 5) = "B.O. Amount"
        .TextMatrix(0, 6) = "Return Stocks"
        .TextMatrix(0, 7) = "R.S. Amount"
        .TextMatrix(0, 8) = "Sold Cases"
        .TextMatrix(0, 9) = "Sold Boxes"
        .TextMatrix(0, 10) = "Sold Pieces"
        .TextMatrix(0, 11) = "Qty Sold"
        .TextMatrix(0, 12) = "Amount"
        .TextMatrix(0, 13) = "Unsold Qty"
        .TextMatrix(0, 14) = "Actual Van Inv."
        .TextMatrix(0, 15) = "Short"
        .TextMatrix(0, 16) = "Over"
        .TextMatrix(0, 17) = "ProductFK"
        .TextMatrix(0, 18) = "TQtyLoad"
        .TextMatrix(0, 19) = "VanInv"
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
        .ColAlignment(13) = vbLeftJustify
        .ColAlignment(14) = vbLeftJustify
        .ColAlignment(15) = vbLeftJustify
        .ColAlignment(16) = vbLeftJustify
        .ColAlignment(16) = vbLeftJustify
        .ColAlignment(17) = vbLeftJustify
        .ColAlignment(18) = vbLeftJustify
        .ColAlignment(19) = vbLeftJustify
    End With
End Sub

Private Sub ResetEntry()
    nsdProduct.ResetValue
    
    txtEntry(0).Text = "0"
    txtEntry(1).Text = "0"
    txtEntry(2).Text = "0"
    
    txtReadOnly(0).Text = "0.00"
    txtReadOnly(1).Text = "0.00"
    txtReadOnly(8).Text = "0"
    txtReadOnly(9).Text = "0"
    
    txtProdDesc.Text = ""
    txtUC.Text = "0.00"
    
    txtReadOnly(2).Text = "0"
    txtReadOnly(3).Text = "0"
    txtReadOnly(4).Text = "0"
    txtReadOnly(5).Text = "0"
    txtReadOnly(6).Text = "0.00"
    txtReadOnly(7).Text = "0"
    txtReadOnly(8).Text = "0"
    
    txtReadOnly(2).Tag = "0"
    txtReadOnly(3).Tag = "0"
    txtReadOnly(4).Tag = "0"
    txtReadOnly(5).Tag = "0"
    txtReadOnly(6).Tag = "0"
    txtReadOnly(7).Tag = "0"
    txtReadOnly(8).Tag = "0"
    
    load_qty = 0
    load_vanqty = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        frmVanInventory.RefreshRecords
        MAIN.UpdateInfoMsg
    End If
    
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

Private Sub nsdProduct_Change()
    txtReadOnly(0).Text = "0.00"
    txtReadOnly(1).Text = "0.00"
    txtReadOnly(8).Text = "0"
    txtReadOnly(9).Text = "0"
    
    txtUC.Text = toMoney(toNumber(nsdProduct.getSelValueAt(3)))
    txtProdDesc.Text = nsdProduct.getSelValueAt(2)
    
    txtReadOnly(2).Text = toNumber(nsdProduct.getSelValueAt(4))
    txtReadOnly(3).Text = toNumber(nsdProduct.getSelValueAt(5))
    txtReadOnly(4).Text = toNumber(nsdProduct.getSelValueAt(6))
    txtReadOnly(5).Text = toNumber(nsdProduct.getSelValueAt(7))
    txtReadOnly(6).Text = toMoney(toNumber(nsdProduct.getSelValueAt(8)))
    If toNumber(nsdProduct.getSelValueAt(9)) = 0 And toNumber(nsdProduct.getSelValueAt(7)) = 0 Then
        txtReadOnly(7).Text = toNumber(nsdProduct.getSelValueAt(12))
        txtReadOnly(7).Tag = toNumber(nsdProduct.getSelValueAt(12))
    Else
        txtReadOnly(7).Text = toNumber(nsdProduct.getSelValueAt(9))
        txtReadOnly(7).Tag = toNumber(nsdProduct.getSelValueAt(9))
    End If
    
    
    load_qty = toNumber(nsdProduct.getSelValueAt(10))
    load_vanqty = toNumber(nsdProduct.getSelValueAt(11))
    
    txtReadOnly(8).Text = toNumber(nsdProduct.getSelValueAt(9))
    
    txtReadOnly(2).Tag = toNumber(nsdProduct.getSelValueAt(4))
    txtReadOnly(3).Tag = toNumber(nsdProduct.getSelValueAt(5))
    txtReadOnly(4).Tag = toNumber(nsdProduct.getSelValueAt(6))
    txtReadOnly(5).Tag = toNumber(nsdProduct.getSelValueAt(7))
    txtReadOnly(6).Tag = toNumber(nsdProduct.getSelValueAt(8))
    txtReadOnly(8).Tag = toNumber(nsdProduct.getSelValueAt(9))
End Sub

Private Sub txtcInvAmount_GotFocus()
    HLText txtcInvAmount
End Sub

Private Sub txtDate_GotFocus()
    HLText txtDate
End Sub

Private Sub txtEntry_Change(Index As Integer)
    If Index = 0 Or Index = 1 Then
        txtReadOnly(7).Text = toNumber(toNumber(txtReadOnly(7).Tag) - toNumber(txtEntry(0).Text) + toNumber(txtEntry(1).Text))
        txtReadOnly(0).Text = toMoney(toNumber(txtEntry(0).Text) * toNumber(txtUC.Text))
        txtReadOnly(1).Text = toMoney(toNumber(txtEntry(1).Text) * toNumber(txtUC.Text))
    ElseIf Index = 2 Then
        
        txtReadOnly(8).Text = toNumber(toNumber(txtReadOnly(7).Text) - toNumber(txtEntry(2).Text), True)
        txtReadOnly(9).Text = toNumber(toNumber(txtEntry(2).Text) - toNumber(txtReadOnly(7).Text), True)
    End If
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index < 3 Then
        KeyAscii = isNumber(KeyAscii)
    End If
End Sub

Private Sub txtEntry_Validate(Index As Integer, Cancel As Boolean)
    If Index < 3 Then
        txtEntry(Index).Text = toNumber(txtEntry(Index).Text)
    End If
End Sub

Private Sub txtReadOnly_Change(Index As Integer)
    Select Case Index
        Case 0
        Case 1
        Case 7: txtEntry_Change 2
    End Select
End Sub

Private Sub txtReadOnly_GotFocus(Index As Integer)
    HLText txtReadOnly(Index)
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

'Procedure used to reset fields
Private Sub ResetFields()
    clearText Me
    InitGrid
    
    dtpDate.Value = Date
    
    ResetEntry
    
    cInvAmount = 0

    txtcInvAmount.Text = "0.00"
       
    dtpDate.SetFocus
End Sub

'Used to display record
Private Sub DisplayForViewing()
    txtVInvNo.Text = rs![VanInventoryNo]
    txtDate.Text = Format$(rs![Date], "MMM-dd-yyyy")
    txtVan.Text = rs![VanName]
    txtcInvAmount.Text = toMoney(rs![SoldAmount])
    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_IC_VanInvDetails WHERE VanInvFK=" & PK & " ORDER BY PK ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            With Grid
                If .Rows = 2 And .TextMatrix(1, 11) = "" Then
                    .TextMatrix(1, 1) = RSDetails![ProductCode]
                    .TextMatrix(1, 2) = RSDetails![Description]
                    .TextMatrix(1, 3) = toMoney(RSDetails![UnitCost(Each)])
                    .TextMatrix(1, 4) = RSDetails![BO]
                    .TextMatrix(1, 5) = toMoney(RSDetails![BOAmount])
                    .TextMatrix(1, 6) = RSDetails![ReturnStock]
                    .TextMatrix(1, 7) = toMoney(RSDetails![ReturnStockAmount])
                    .TextMatrix(1, 8) = RSDetails![SoldCases]
                    .TextMatrix(1, 9) = RSDetails![SoldBoxes]
                    .TextMatrix(1, 10) = RSDetails![SoldPieces]
                    .TextMatrix(1, 11) = RSDetails![TotalQty]
                    .TextMatrix(1, 12) = toMoney(RSDetails![Amount])
                    .TextMatrix(1, 13) = RSDetails![TotalUnsoldQty]
                    .TextMatrix(1, 14) = RSDetails![ActualVanInv]
                    .TextMatrix(1, 15) = RSDetails![Short]
                    .TextMatrix(1, 16) = RSDetails![Over]
                    .TextMatrix(1, 17) = RSDetails![ProductFK]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![ProductCode]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Description]
                    .TextMatrix(.Rows - 1, 3) = toMoney(RSDetails![UnitCost(Each)])
                    .TextMatrix(.Rows - 1, 4) = RSDetails![BO]
                    .TextMatrix(.Rows - 1, 5) = toMoney(RSDetails![BOAmount])
                    .TextMatrix(.Rows - 1, 6) = RSDetails![ReturnStock]
                    .TextMatrix(.Rows - 1, 7) = toMoney(RSDetails![ReturnStockAmount])
                    .TextMatrix(.Rows - 1, 8) = RSDetails![SoldCases]
                    .TextMatrix(.Rows - 1, 9) = RSDetails![SoldBoxes]
                    .TextMatrix(.Rows - 1, 10) = RSDetails![SoldPieces]
                    .TextMatrix(.Rows - 1, 11) = RSDetails![TotalQty]
                    .TextMatrix(.Rows - 1, 12) = toMoney(RSDetails![Amount])
                    .TextMatrix(.Rows - 1, 13) = RSDetails![TotalUnsoldQty]
                    .TextMatrix(.Rows - 1, 14) = RSDetails![ActualVanInv]
                    .TextMatrix(.Rows - 1, 15) = RSDetails![Short]
                    .TextMatrix(.Rows - 1, 16) = RSDetails![Over]
                    .TextMatrix(.Rows - 1, 17) = RSDetails![ProductFK]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 17
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 2
        End If
    End If

    RSDetails.Close

    'Disable commands
    LockInput Me, True

    dtpDate.Visible = False
    txtDate.Visible = True
    picProd.Visible = False
    cmdSave.Visible = False
    btnAddInv.Visible = False

    'Resize and reposition the controls
    Shape3.Top = 900
    Label11.Top = 900
    Grid.Top = 1200
    Grid.Height = 3565
    txtcInvAmount.Top = txtcInvAmount.Top - 2000
    Label9.Top = Label9.Top - 2000
    cmdUsrHistory.Top = cmdUsrHistory.Top - 2000
    btnAddInv.Top = btnAddInv.Top - 2000
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

Private Sub InitNSD()
    'For Product
    With nsdProduct
        .ClearColumn
        .AddColumn "Product Code", 2064.882
        .AddColumn "Description", 4085.26
        .AddColumn "Unit Cost(Each)", 1500
        .AddColumn "Sold Cases", 1500
        .AddColumn "Sold Boxes", 1500
        .AddColumn "Sold Pieces", 1500
        .AddColumn "Total Qty Sold", 1500
        .AddColumn "Amount", 1500
        .AddColumn "Qty Unsold", 1500
        .AddColumn "Qty Load", 1500
        .AddColumn "Van Inv.", 1500
        .AddColumn "Total Qty Load", 1500
        
        .Connection = CN.ConnectionString
        
        .sqlFields = "ProductCode,Description,UnitCost,SoldCases,SoldBoxes,SoldPieces,TotalQtySold,Amount,TotalUnsoldQty,QtyLoad,VanInv,TotalQtyLoad,ProductFK,LLFK,PK"
        .sqlTables = "qry_AR_LoadingDetailsInfo"
        .sqlSortOrder = "PK ASC"
        .sqlwCondition = "LLFK = " & LLFK
        .BoundField = "ProductFK"
        .PageBy = 25
        .DisplayCol = 1
        
        .setDropWindowSize 6800, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Sold Products From " & LLDate
        
    End With

End Sub
