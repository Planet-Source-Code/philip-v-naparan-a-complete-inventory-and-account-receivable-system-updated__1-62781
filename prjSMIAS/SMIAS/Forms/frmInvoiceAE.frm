VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmInvoiceAE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8715
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
   Icon            =   "frmInvoiceAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtLess 
      Height          =   285
      Left            =   9450
      TabIndex        =   28
      Text            =   "0.00"
      Top             =   6075
      Width           =   1500
   End
   Begin VB.CommandButton cmdPH 
      Caption         =   "Payment History"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2100
      TabIndex        =   66
      Top             =   7875
      Width           =   1590
   End
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -75
      TabIndex        =   6
      Top             =   7725
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   53
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9450
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "0.00"
      Top             =   6375
      Width           =   1500
   End
   Begin VB.TextBox txtEntry 
      Height          =   990
      Index           =   8
      Left            =   225
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   27
      Tag             =   "Remarks"
      Top             =   6300
      Width           =   5805
   End
   Begin VB.TextBox txtVan 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7875
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   75
      Width           =   3075
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
      TabIndex        =   30
      Text            =   "0.00"
      Top             =   6675
      Width           =   1500
   End
   Begin VB.TextBox txtAP 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9450
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "0.00"
      Top             =   6975
      Width           =   1500
   End
   Begin VB.PictureBox picCusInfo 
      BorderStyle     =   0  'None
      Height          =   1740
      Left            =   225
      ScaleHeight     =   1740
      ScaleWidth      =   10740
      TabIndex        =   49
      Top             =   1050
      Width           =   10740
      Begin VB.TextBox txtDP 
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7725
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   1350
         Width           =   1500
      End
      Begin VB.CommandButton cmdReset 
         Height          =   315
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Reset Selection"
         Top             =   0
         Width           =   315
      End
      Begin VB.ComboBox cbBI 
         Height          =   315
         ItemData        =   "frmInvoiceAE.frx":038A
         Left            =   7725
         List            =   "frmInvoiceAE.frx":0394
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   975
         Width           =   2565
      End
      Begin VB.ComboBox cbPT 
         Height          =   315
         ItemData        =   "frmInvoiceAE.frx":03B3
         Left            =   1275
         List            =   "frmInvoiceAE.frx":03C0
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1350
         Width           =   2565
      End
      Begin VB.ComboBox cbCA 
         Height          =   315
         ItemData        =   "frmInvoiceAE.frx":03EB
         Left            =   1275
         List            =   "frmInvoiceAE.frx":03F5
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   975
         Width           =   2565
      End
      Begin VB.TextBox txtCusAdd 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1275
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   375
         Width           =   4425
      End
      Begin VB.TextBox txtCusCP 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7725
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   0
         Width           =   3000
      End
      Begin VB.CommandButton cmdNew 
         Height          =   315
         Left            =   5025
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Create New"
         Top             =   0
         Width           =   315
      End
      Begin ctrlNSDataCombo.NSDataCombo nsdCustomer 
         Height          =   315
         Left            =   1275
         TabIndex        =   7
         Top             =   0
         Width           =   3690
         _ExtentX        =   6509
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Down Payment"
         ForeColor       =   &H0000011D&
         Height          =   240
         Left            =   5625
         TabIndex        =   65
         Top             =   1350
         Width           =   2040
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Payment Type"
         Height          =   240
         Index           =   12
         Left            =   -975
         TabIndex        =   55
         Top             =   1350
         Width           =   2190
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Billed In   Full Payment"
         Height          =   240
         Index           =   7
         Left            =   5550
         TabIndex        =   54
         Top             =   975
         Width           =   3165
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Charge Account"
         Height          =   240
         Index           =   6
         Left            =   -975
         TabIndex        =   53
         Top             =   975
         Width           =   2190
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Address"
         Height          =   240
         Index           =   5
         Left            =   -975
         TabIndex        =   52
         Top             =   375
         Width           =   2190
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Contact Person"
         Height          =   240
         Index           =   3
         Left            =   5475
         TabIndex        =   51
         Top             =   0
         Width           =   2190
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Sold To"
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
         TabIndex        =   50
         Top             =   0
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdGenerate 
      Height          =   315
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Generate"
      Top             =   150
      Width           =   315
   End
   Begin VB.PictureBox picPurchase 
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   225
      ScaleHeight     =   540
      ScaleWidth      =   10740
      TabIndex        =   40
      Top             =   2925
      Width           =   10740
      Begin VB.CheckBox ckFree 
         Height          =   315
         Left            =   9075
         TabIndex        =   63
         Top             =   225
         Width           =   240
      End
      Begin VB.ComboBox cbDisc 
         Height          =   315
         Left            =   6750
         TabIndex        =   22
         Text            =   "0"
         Top             =   225
         Width           =   765
      End
      Begin VB.TextBox txtAmount 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   7575
         Locked          =   -1  'True
         TabIndex        =   23
         Text            =   "0.00"
         Top             =   225
         Width           =   1410
      End
      Begin VB.TextBox txtTQty 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   5850
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "0"
         Top             =   225
         Width           =   825
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   4
         Left            =   5175
         TabIndex        =   20
         Text            =   "0"
         Top             =   225
         Width           =   615
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   3
         Left            =   4575
         TabIndex        =   19
         Text            =   "0"
         Top             =   225
         Width           =   540
      End
      Begin VB.TextBox txtSP 
         Height          =   285
         Left            =   2700
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   225
         Width           =   1185
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   2
         Left            =   3975
         TabIndex        =   18
         Text            =   "0"
         Top             =   225
         Width           =   540
      End
      Begin VB.CommandButton btnSold 
         Caption         =   "Sold"
         Enabled         =   0   'False
         Height          =   315
         Left            =   9900
         TabIndex        =   24
         Top             =   225
         Width           =   840
      End
      Begin ctrlNSDataCombo.NSDataCombo nsdProduct 
         Height          =   315
         Left            =   0
         TabIndex        =   16
         Top             =   225
         Width           =   2640
         _ExtentX        =   4657
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
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "FREE"
         Height          =   240
         Index           =   19
         Left            =   9075
         TabIndex        =   64
         Top             =   0
         Width           =   540
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Disc.%"
         Height          =   240
         Index           =   14
         Left            =   6750
         TabIndex        =   61
         Top             =   0
         Width           =   840
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   240
         Index           =   17
         Left            =   7575
         TabIndex        =   47
         Top             =   0
         Width           =   540
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Qty"
         Height          =   240
         Index           =   16
         Left            =   5850
         TabIndex        =   46
         Top             =   0
         Width           =   840
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Pieces"
         Height          =   240
         Index           =   15
         Left            =   5175
         TabIndex        =   45
         Top             =   0
         Width           =   540
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Boxes"
         Height          =   240
         Index           =   11
         Left            =   4575
         TabIndex        =   44
         Top             =   0
         Width           =   540
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Cases"
         Height          =   240
         Index           =   10
         Left            =   3975
         TabIndex        =   43
         Top             =   0
         Width           =   540
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Price(Each)"
         Height          =   240
         Index           =   9
         Left            =   2700
         TabIndex        =   42
         Top             =   0
         Width           =   1290
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Product Code"
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
         Index           =   8
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   1515
      End
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   300
      Picture         =   "frmInvoiceAE.frx":0407
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Remove"
      Top             =   3975
      Visible         =   0   'False
      Width           =   275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   8175
      TabIndex        =   34
      Top             =   7875
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   9615
      TabIndex        =   35
      Top             =   7875
      Width           =   1335
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   225
      TabIndex        =   33
      Top             =   7875
      Width           =   1755
   End
   Begin VB.TextBox txtBal 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9450
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "0.00"
      Top             =   7275
      Width           =   1500
   End
   Begin VB.TextBox txtEntry 
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1425
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   2115
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2190
      Left            =   225
      TabIndex        =   25
      Top             =   3825
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
      TabIndex        =   2
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
      TabIndex        =   3
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
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   5325
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Less"
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   7350
      TabIndex        =   67
      Top             =   6075
      Width           =   2040
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Discount"
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   7350
      TabIndex        =   62
      Top             =   6375
      Width           =   2040
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Remarks"
      Height          =   240
      Index           =   4
      Left            =   -150
      TabIndex        =   60
      Top             =   6075
      Width           =   990
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Van"
      Height          =   240
      Index           =   13
      Left            =   6600
      TabIndex        =   58
      Top             =   75
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Amount"
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
      TabIndex        =   57
      Top             =   6675
      Width           =   2040
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Amount Paid"
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   7350
      TabIndex        =   56
      Top             =   6975
      Width           =   2040
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Salesman"
      Height          =   240
      Index           =   18
      Left            =   6600
      TabIndex        =   48
      Top             =   450
      Width           =   1215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   225
      X2              =   10950
      Y1              =   2850
      Y2              =   2850
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   225
      X2              =   10950
      Y1              =   2850
      Y2              =   2850
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Sold Products"
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
      TabIndex        =   39
      Top             =   3525
      Width           =   4365
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Balance"
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   7350
      TabIndex        =   38
      Top             =   7275
      Width           =   2040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   225
      X2              =   10950
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   0
      X1              =   225
      X2              =   10950
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   " Date"
      Height          =   240
      Index           =   1
      Left            =   150
      TabIndex        =   37
      Top             =   525
      Width           =   1215
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
      Index           =   0
      Left            =   450
      TabIndex        =   36
      Top             =   150
      Width           =   915
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   225
      Top             =   3525
      Width           =   10740
   End
End
Attribute VB_Name = "frmInvoiceAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public LLFK                 As Long 'Last Loading FK
Public LLVFK                As Long 'Last Loading Van FK
Public LLDate               As String
Public CloseMe              As Boolean
Public ForCusAcc            As Boolean

Dim PCase                   As Long 'Pieces per case
Dim PBox                    As Long 'Pieces per box

Dim old_pieces              As Long 'Old pieces value
Dim old_boxes               As Long 'Old boxes value
Dim old_cases               As Long 'Old cases value

Dim cIAmount                As Currency 'Current Invoice Amount
Dim cDAmount                As Currency 'Current Invoice Discount Amount
Dim cIRowCount              As Integer

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset 'Main recordset for Invoice

Private Sub btnSold_Click()
    If nsdProduct.Text = "" Then nsdProduct.SetFocus: Exit Sub

    If toNumber(txtSP.Text) <= 0 Then
        MsgBox "Please enter a valid sales price.", vbExclamation
        txtSP.SetFocus
        Exit Sub
    End If

    Dim CurrRow As Integer

    CurrRow = getFlexPos(Grid, 11, nsdProduct.BoundText)

    'Add to grid
    With Grid
        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 11) = "" Then
                .TextMatrix(1, 1) = nsdProduct.Text
                .TextMatrix(1, 2) = nsdProduct.getSelValueAt(2)
                .TextMatrix(1, 3) = txtSP.Text 'Unit price
                .TextMatrix(1, 4) = txtEntry(2).Text
                .TextMatrix(1, 5) = txtEntry(3).Text
                .TextMatrix(1, 6) = txtEntry(4).Text
                .TextMatrix(1, 7) = txtTQty.Text
                .TextMatrix(1, 8) = toNumber(cbDisc.Text)
                .TextMatrix(1, 9) = txtAmount.Text
                .TextMatrix(1, 10) = changeYNValue(ckFree.Value)
                .TextMatrix(1, 11) = nsdProduct.BoundText
                .TextMatrix(1, 12) = toNumber(toNumber(cbDisc.Text) / 100) * toNumber(toNumber(txtTQty.Text) * toNumber(txtSP.Text))
                .TextMatrix(1, 13) = txtSP.Tag 'Unit cost
            Else
ADD_NEW_HERE:
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = nsdProduct.Text
                .TextMatrix(.Rows - 1, 2) = nsdProduct.getSelValueAt(2)
                .TextMatrix(.Rows - 1, 3) = txtSP.Text 'Unit price
                .TextMatrix(.Rows - 1, 4) = txtEntry(2).Text
                .TextMatrix(.Rows - 1, 5) = txtEntry(3).Text
                .TextMatrix(.Rows - 1, 6) = txtEntry(4).Text
                .TextMatrix(.Rows - 1, 7) = txtTQty.Text
                .TextMatrix(.Rows - 1, 8) = toNumber(cbDisc.Text)
                .TextMatrix(.Rows - 1, 9) = txtAmount.Text
                .TextMatrix(.Rows - 1, 10) = changeYNValue(ckFree.Value)
                .TextMatrix(.Rows - 1, 11) = nsdProduct.BoundText
                .TextMatrix(.Rows - 1, 12) = toNumber(toNumber(cbDisc.Text) / 100) * toNumber(toNumber(txtTQty.Text) * toNumber(txtSP.Text))
                .TextMatrix(.Rows - 1, 13) = txtSP.Tag 'Unit cost
                
                .Row = .Rows - 1
            End If
            'Increase the record count
            cIRowCount = cIRowCount + 1
        Else
            'If free option is not equal or discount is not equal or sales price is not equal then add new sold item
            If .TextMatrix(CurrRow, 10) <> changeYNValue(ckFree.Value) Or toNumber(.TextMatrix(CurrRow, 8)) <> toNumber(cbDisc.Text) Or toNumber(.TextMatrix(CurrRow, 3)) <> toNumber(txtSP.Text) Then
                GoTo ADD_NEW_HERE
            End If
            
            If MsgBox("Invoice payment already exist.Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
                .Row = CurrRow
                
                'Restore back the invoice amount and discount
                cIAmount = cIAmount - toNumber(Grid.TextMatrix(.RowSel, 9))
                txtTA.Text = Format$(cIAmount, "#,##0.00")
                cDAmount = cDAmount - toNumber(Grid.TextMatrix(.RowSel, 12))
                txtDesc.Text = Format$(cDAmount, "#,##0.00")
                
                .TextMatrix(CurrRow, 1) = nsdProduct.Text
                .TextMatrix(CurrRow, 2) = nsdProduct.getSelValueAt(2)
                .TextMatrix(CurrRow, 3) = txtSP.Text 'Unit price
                .TextMatrix(CurrRow, 4) = txtEntry(2).Text
                .TextMatrix(CurrRow, 5) = txtEntry(3).Text
                .TextMatrix(CurrRow, 6) = txtEntry(4).Text
                .TextMatrix(CurrRow, 7) = txtTQty.Text
                .TextMatrix(CurrRow, 8) = toNumber(cbDisc.Text)
                .TextMatrix(CurrRow, 9) = txtAmount.Text
                .TextMatrix(CurrRow, 10) = changeYNValue(ckFree.Value)
                .TextMatrix(CurrRow, 11) = nsdProduct.BoundText
                .TextMatrix(CurrRow, 12) = toNumber(toNumber(cbDisc.Text) / 100) * toNumber(toNumber(txtTQty.Text) * toNumber(txtSP.Text))
                .TextMatrix(CurrRow, 13) = txtSP.Tag 'Unit cost
            Else
                Exit Sub
            End If
        End If
        'Add the amount to current load amount
        cIAmount = cIAmount + toNumber(txtAmount.Text)
        cDAmount = cDAmount + toNumber(toNumber(cbDisc.Text) / 100) * (toNumber(toNumber(txtTQty.Text) * toNumber(txtSP.Text)))
        txtDesc.Text = Format$(cDAmount, "#,##0.00")
        txtTA.Text = Format$(cIAmount, "#,##0.00")
        'Highlight the current row's column
        .ColSel = 11
        'Display a remove button
        Grid_Click
        'Reset the entry fields
        ResetEntry
    End With
End Sub

Private Sub btnRemove_Click()
    'Remove selected load product
    With Grid
        'Update amount to current invoice amount
        cIAmount = cIAmount - toNumber(Grid.TextMatrix(.RowSel, 9))
        txtTA.Text = Format$(cIAmount, "#,##0.00")
        'Update discount to current invoice disc
        cDAmount = cDAmount - toNumber(Grid.TextMatrix(.RowSel, 12))
        txtDesc.Text = Format$(cDAmount, "#,##0.00")
        'Update the record count
        cIRowCount = cIRowCount - 1
        
        If .Rows = 2 Then Grid.Rows = Grid.Rows + 1
        .RemoveItem (.RowSel)
    End With

    btnRemove.Visible = False
    Grid_Click
    
End Sub

Private Sub cbBI_Click()
    'Not paid Option
    If cbBI.ListIndex = 0 Then
        txtDP.Enabled = False
        txtDP.Text = "0.00"
        cbPT.ListIndex = -1
        cbPT.Enabled = False
    Else 'If Partial
        txtDP.Enabled = True
        cbPT.ListIndex = 0
        cbPT.Enabled = True
    End If
End Sub

Private Sub cbCA_Click()
    txtDP.Text = "0.00"
    'Charge Account Option
    If cbCA.ListIndex = 1 Then 'If Credit
        cbBI.Visible = True
        Label2.Visible = True
        txtDP.Visible = True
        cbPT.ListIndex = -1
        cbPT.Enabled = False
        
        Label1.Visible = True
        Label9.Visible = True
        txtAP.Visible = True
        txtBal.Visible = True
        
    Else 'If Cash
        cbBI.Visible = False
        Label2.Visible = False
        txtDP.Visible = False
        cbPT.ListIndex = 0
        cbPT.Enabled = True
        
        Label1.Visible = False
        Label9.Visible = False
        txtAP.Visible = False
        txtBal.Visible = False
    End If
End Sub

Private Sub cbDisc_Change()
    txtTQty_Change
End Sub

Private Sub cbDisc_Click()
    txtTQty_Change
End Sub

Private Sub ckFree_Click()
If ckFree.Value = 1 Then 'If checked
    cbDisc.Text = "0"
    cbDisc.Visible = False
    txtAmount.Text = "0.00"
    txtAmount.Visible = False
    Labels(14).Visible = False
    Labels(17).Visible = False
Else
    cbDisc.Visible = True
    txtAmount.Visible = True
    Labels(14).Visible = True
    Labels(17).Visible = True
End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cbDisc_Validate(Cancel As Boolean)
    cbDisc.Text = toNumber(cbDisc.Text)
End Sub

Private Sub cmdGenerate_Click()
    GeneratePK
End Sub

Private Sub cmdNew_Click()
    With frmCustomerAE
        .State = adStatePopupMode
        Set .srcText = txtNCus
        Set .srcTextAdd = txtCusAdd
        Set .srcTextCP = txtCusCP
        Set .srcTextDisc = cbDisc
        .show vbModal
    End With
    If txtNCus.Tag = "" And txtNCus.Text = "" Then Exit Sub
    nsdCustomer.DisableDropdown = True
    nsdCustomer.Text = txtNCus.Text
End Sub

Private Sub cmdPH_Click()
    frmInvoiceViewerPH.INV_PK = PK
    frmInvoiceViewerPH.Caption = "Payment History Viewer"
    frmInvoiceViewerPH.lblTitle.Caption = "Payment History Viewer"
    frmInvoiceViewerPH.show vbModal
End Sub

Private Sub cmdReset_Click()
    With nsdCustomer
        .ResetValue
        .DisableDropdown = False
    End With
    
    txtNCus.Tag = ""
    txtNCus.Text = ""
    
    txtCusAdd.Text = ""
    txtCusCP.Text = ""
End Sub

Private Sub cmdSave_Click()
    'Verify the entries
    If txtEntry(0).Text = "" Then
        MsgBox "Please enter an invoice number.", vbExclamation
        txtEntry(0).SetFocus
        Exit Sub
    End If
       
    If dcSalesman.BoundText = "" Then
        MsgBox "Please select a salesman in the list.", vbExclamation
        dcSalesman.SetFocus
        Exit Sub
    End If
    
    If nsdCustomer.BoundText = "" And txtNCus.Tag = "" Then
        MsgBox "Please select a customer.", vbExclamation
        nsdCustomer.SetFocus
        Exit Sub
    End If
    
    If txtBal.Visible = True And toNumber(txtBal.Text) <= 0 Then
            MsgBox "Please enter a valid downpayment.", vbExclamation
            txtDP.SetFocus
            Exit Sub
    End If
    
    If cIRowCount < 1 Then
        MsgBox "Please enter a sold product first before you can save this record.", vbExclamation
        nsdProduct.SetFocus
        Exit Sub
    End If
    
    If isRecordExist("tbl_AR_Invoice", "InvoiceNo", txtEntry(0).Text, True) = True Then
        MsgBox "Invoice No. already exist.Please change it.", vbExclamation
        txtEntry(0).SetFocus
        Exit Sub
    End If
    

    If MsgBox("This save the record.Do you want to proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub


    Dim RSDetails As New Recordset


    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM tbl_AR_InvoiceDetails WHERE InvoiceFK=" & PK, CN, adOpenStatic, adLockOptimistic

    Screen.MousePointer = vbHourglass

    Dim c As Integer

    On Error GoTo err

    CN.BeginTrans

    If cbPT.ListIndex = 2 Then
        Screen.MousePointer = vbDefault
        
        frmPDCManagerAE.State = adStatePopupMode
        If toNumber(txtDP.Text) > 0 Then
            frmPDCManagerAE.txtEntry(6).Text = "Downpayment for Invoice No. " & txtEntry(0).Text & "."
            frmPDCManagerAE.txtEntry(3).Text = toMoney(toNumber(txtDP.Text))
        Else
            frmPDCManagerAE.txtEntry(6).Text = "Payment for Invoice No. " & txtEntry(0).Text & "."
            frmPDCManagerAE.txtEntry(3).Text = toMoney(toNumber(txtTA.Text))
        End If
        frmPDCManagerAE.show vbModal
        
        Screen.MousePointer = vbHourglass
    End If

    If nsdCustomer.BoundText <> "" Then
        If toNumber(getRecordCount("tbl_AR_Invoice", "WHERE SoldToPK =" & nsdCustomer.BoundText)) >= 1 And getValueAt("SELECT PK,Status FROM tbl_AR_Customer WHERE PK=" & nsdCustomer.BoundText, "Status") = "New" Then
            ChangeValue CN, "tbl_AR_Customer", "Status", "Old", False, "WHERE PK=" & nsdCustomer.BoundText
        End If
    Else
        If toNumber(getRecordCount("tbl_AR_Invoice", "WHERE SoldToPK =" & txtNCus.Tag)) >= 1 And getValueAt("SELECT PK,Status FROM tbl_AR_Customer WHERE PK=" & txtNCus.Tag, "Status") = "New" Then
            ChangeValue CN, "tbl_AR_Customer", "Status", "Old", False, "WHERE PK=" & txtNCus.Tag
        End If
    End If


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
        ![InvoiceNo] = txtEntry(0).Text
        ![Date] = dtpDate.Value
        If nsdCustomer.BoundText <> "" Then
            ![SoldToPK] = nsdCustomer.BoundText
        Else
            ![SoldToPK] = txtNCus.Tag
        End If
        ![VanFK] = LLVFK
        ![SalesmanFK] = dcSalesman.BoundText
        ![LastLoadingFK] = LLFK
        ![ChargeAccount] = cbCA.Text
        ![PaymentType] = cbPT.Text
        If cbBI.Visible = True Then
            ![BilledIn] = cbBI.Text
            ![AmountPaid] = toNumber(txtAP.Text)
            ![Paid] = "N"
        Else
            ![BilledIn] = "Full Payment"
            ![AmountPaid] = toNumber(txtTA.Text)
            ![Paid] = "Y"
        End If
        ![DownPayment] = toNumber(txtDP.Text)
        ![DAmount] = cDAmount
        ![TAmount] = cIAmount - toNumber(txtLess.Text)
        ![Less] = toNumber(txtLess.Text)
        ![Remarks] = txtEntry(8).Text
        
        .Update
        
    End With

    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then
            
                RSDetails.AddNew

                RSDetails![PK] = getIndex("tbl_AR_InvoiceDetails")

                RSDetails![InvoiceFK] = PK
                RSDetails![ProductFK] = toNumber(.TextMatrix(c, 11))
                RSDetails![UnitCost(Each)] = toNumber(.TextMatrix(c, 13))
                RSDetails![SalesPrice(Each)] = toNumber(.TextMatrix(c, 3))
                
                RSDetails![SoldCases] = toNumber(.TextMatrix(c, 4))
                RSDetails![SoldBoxes] = toNumber(.TextMatrix(c, 5))
                RSDetails![SoldPieces] = toNumber(.TextMatrix(c, 6))
                RSDetails![TotalQty] = toNumber(.TextMatrix(c, 7))
                RSDetails![Disc] = toNumber(.TextMatrix(c, 8))
                RSDetails![Discount] = toNumber(.TextMatrix(c, 12))
                RSDetails![Free] = .TextMatrix(c, 10)

                RSDetails.Update

            End If

        Next c
    End With

    'Clear variables
    c = 0
    Set RSDetails = Nothing

    CN.CommitTrans

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

Private Sub Form_Activate()
    On Error Resume Next
    If CloseMe = True Then
        Unload Me
    Else
        txtEntry(0).SetFocus
    End If
End Sub

Private Sub Form_Load()
    InitGrid

    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        frmInvoiceAEPickFrom.show vbModal
        
        'Bind the data combo
        bind_dc "SELECT * FROM tbl_AR_Salesman", "Name", dcSalesman, "PK", True
        'Initialize controls
        cbBI.ListIndex = 0
        cbCA.ListIndex = 0
        cbPT.ListIndex = 0
        InitNSD
        

        'Discount per product
        Labels(14).Visible = True
        cbDisc.Visible = True
        
        'Set the recordset
         rs.Open "SELECT * FROM tbl_AR_Invoice WHERE PK=" & PK, CN, adOpenStatic, adLockOptimistic
         dtpDate.Value = Date
         Caption = "Create New Entry"
         cmdUsrHistory.Enabled = False
         GeneratePK
    Else
        Screen.MousePointer = vbHourglass
        'Set the recordset
        rs.Open "SELECT * FROM qry_AR_Invoice WHERE PK=" & PK, CN, adOpenStatic, adLockOptimistic
        
        If rs![Paid] = "Y" Then
            Caption = "View Record (Paid)"
        Else
            Caption = "View Record (Not Paid)"
        End If
        
        cmdCancel.Caption = "Close"
        cmdUsrHistory.Enabled = True
        
        txtEntry(0).Width = txtDate.Width
        
        DisplayForViewing
        If cbCA.ListIndex = 1 Then cmdPH.Enabled = True
        
        If ForCusAcc = True Then
            Me.Icon = frmAccCustomer.Icon
        Else
            
            MsgBox "This is use for viewing the record only." & vbCrLf & _
               "You cannot perform any changes in this form." & vbCrLf & vbCrLf & _
               "Note:If you have mistake in adding this record then " & vbCrLf & _
               "void this record and re-enter.", vbExclamation
        End If

        Screen.MousePointer = vbDefault
    End If
    
    'Initialize Graphics
    With MAIN
        cmdGenerate.Picture = .i16x16.ListImages(14).Picture
        cmdNew.Picture = .i16x16.ListImages(10).Picture
        cmdReset.Picture = .i16x16.ListImages(15).Picture
    End With
    
    'Fill the discount combo
    cbDisc.AddItem "0.01"
    cbDisc.AddItem "0.02"
    cbDisc.AddItem "0.03"
    cbDisc.AddItem "0.04"
    cbDisc.AddItem "0.05"
    cbDisc.AddItem "0.06"
    cbDisc.AddItem "0.07"
    cbDisc.AddItem "0.08"
    cbDisc.AddItem "0.09"
    cbDisc.AddItem "0.1"
    
End Sub

'Procedure used to generate PK
Private Sub GeneratePK()
    PK = getIndex("tbl_AR_Invoice")
    txtEntry(0).Text = "INV" & GenerateID(PK, Format$(Date, "yyyy") & Format$(Date, "mm") & Format$(Date, "dd") & "-", "0")
End Sub

'Procedure used to initialize the grid
Private Sub InitGrid()
    cIRowCount = 0
    With Grid
        .Clear
        .ClearStructure
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .Cols = 14
        .ColSel = 11
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 2025
        .ColWidth(2) = 2505
        .ColWidth(3) = 1545
        .ColWidth(4) = 900
        .ColWidth(5) = 900
        .ColWidth(6) = 900
        .ColWidth(7) = 900
        .ColWidth(8) = 900
        .ColWidth(9) = 1545
        .ColWidth(10) = 750
        .ColWidth(11) = 0
        .ColWidth(12) = 0
        .ColWidth(13) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Product Code"
        .TextMatrix(0, 2) = "Description"
        .TextMatrix(0, 3) = "Sales Price(Each)"
        .TextMatrix(0, 4) = "Cases"
        .TextMatrix(0, 5) = "Boxes"
        .TextMatrix(0, 6) = "Pieces"
        .TextMatrix(0, 7) = "Qty Sold"
        .TextMatrix(0, 8) = "Disc%"
        .TextMatrix(0, 9) = "Amount"
        .TextMatrix(0, 10) = "FREE"
        .TextMatrix(0, 11) = "ProductFK"
        .TextMatrix(0, 12) = "Disc"
        .TextMatrix(0, 13) = "UC" 'Unit Cost
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
    End With
End Sub

Private Sub ResetEntry()
    txtEntry(2).Text = "0"
    txtEntry(3).Text = "0"
    txtEntry(4).Text = "0"

    cbDisc.Text = toNumber(nsdCustomer.getSelValueAt(8))
    ckFree.Value = 0
    
    nsdProduct.ResetValue
    txtSP.Tag = 0
    txtSP.Text = "0.00"
    PCase = 0
    PBox = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        frmInvoice.RefreshRecords
        MAIN.UpdateInfoMsg
    End If
    
    Set frmInvoiceAE = Nothing
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

Private Sub nsdCustomer_Change()
    If nsdCustomer.DisableDropdown = False Then
    
        txtCusAdd.Text = nsdCustomer.getSelValueAt(3)
        
        If nsdCustomer.getSelValueAt(4) <> "" Then txtCusAdd.Text = txtCusAdd.Text & "," & nsdCustomer.getSelValueAt(4)
        If nsdCustomer.getSelValueAt(5) <> "" Then txtCusAdd.Text = txtCusAdd.Text & "," & nsdCustomer.getSelValueAt(5)
        If nsdCustomer.getSelValueAt(6) <> "" Then txtCusAdd.Text = txtCusAdd.Text & "," & nsdCustomer.getSelValueAt(6)
             
        txtCusCP.Text = nsdCustomer.getSelValueAt(7)
        If ckFree.Value = 1 Then
            cbDisc.Text = "0"
            cbDisc.Enabled = False
        Else
            cbDisc.Text = toNumber(nsdCustomer.getSelValueAt(8))
            cbDisc.Enabled = True
        End If
    End If
End Sub

Private Sub nsdProduct_Change()
    txtEntry(2).Text = "0"
    txtEntry(3).Text = "0"
    txtEntry(4).Text = "0"
    
    txtSP.Tag = nsdProduct.getSelValueAt(3) 'Unit Cost
    txtSP.Text = nsdProduct.getSelValueAt(4) 'Selling Price
    PCase = toNumber(nsdProduct.getSelValueAt(6))
    PBox = toNumber(nsdProduct.getSelValueAt(5))
End Sub

Private Sub txtAmount_GotFocus()
    HLText txtAmount
End Sub

Private Sub txtAP_Change()
    txtBal.Text = toMoney(toNumber(txtTA.Text) - toNumber(txtAP.Text))
End Sub

Private Sub txtBal_GotFocus()
    HLText txtBal
End Sub

Private Sub txtDate_GotFocus()
    HLText txtDate
End Sub

Private Sub txtDesc_GotFocus()
    HLText txtDesc
End Sub

Private Sub txtDP_Change()
    txtAP.Text = toMoney(toNumber(txtDP.Text))
End Sub

Private Sub txtDP_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtEntry_Change(Index As Integer)
    If Index > 1 And Index < 5 Then
        txtTQty.Text = (toNumber(txtEntry(2).Text) * PCase) + (toNumber(txtEntry(3).Text) * PBox) + toNumber(txtEntry(4).Text)
    End If
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

Private Sub txtLess_GotFocus()
    HLText txtLess
End Sub

Private Sub txtLess_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtSP_Change()
    txtTQty_Change
End Sub

Private Sub txtSP_Validate(Cancel As Boolean)
    txtSP.Text = toMoney(toNumber(txtSP.Text))
End Sub

Private Sub txtTA_Change()
    txtAP_Change
End Sub

Private Sub txtTQty_Change()
    If toNumber(txtTQty.Text) < 1 Then
        btnSold.Enabled = False
    Else
        btnSold.Enabled = True
    End If
    txtAmount.Text = toMoney((toNumber(txtTQty.Text) * toNumber(txtSP.Text)) - ((toNumber(cbDisc.Text) / 100) * toNumber(toNumber(txtTQty.Text) * toNumber(txtSP.Text))))
End Sub

Private Sub txtTQty_GotFocus()
    HLText txtTQty
End Sub

Private Sub txtSP_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

'Procedure used to reset fields
Private Sub ResetFields()
    InitGrid
    ResetEntry
    cmdReset_Click
    
    dtpDate.Value = Date
    cbCA.ListIndex = 0
    
    txtEntry(8).Text = ""
    
    txtDesc.Text = "0.00"
    txtTA.Text = "0.00"
    txtAP.Text = "0.00"
    txtBal.Text = "0.00"

    cIAmount = 0
    cDAmount = 0

    txtEntry(0).SetFocus
End Sub

'Used to display record
Private Sub DisplayForViewing()
    On Error GoTo err
    txtEntry(0).Text = rs![InvoiceNo]
    txtDate.Text = Format$(rs![Date], "MMM-dd-yyyy")
    txtVan.Text = rs![VanName]
    bind_dc "SELECT * FROM tbl_AR_Salesman", "Name", dcSalesman, "PK", True
    dcSalesman.BoundText = rs![SalesmanFK]
    'Initialize nsd controls
    nsdCustomer.DisableDropdown = True
    nsdCustomer.TextReadOnly = True
    nsdCustomer.Text = rs![SoldTo]
    txtCusAdd.Text = rs![Address]
    txtCusCP.Text = rs![ContactPerson]
    'Display charge account
    If rs![ChargeAccount] = "Cash" Then
        cbCA.ListIndex = 0
    Else
        cbCA.ListIndex = 1
    End If
    'Display payment type
    If rs![PaymentType] = "Cash" Then
        cbPT.ListIndex = 0
    ElseIf rs![PaymentType] = "On Date Check" Then
        cbPT.ListIndex = 1
    Else
        cbPT.ListIndex = 2
    End If
    'Display billed in
    If rs![BilledIn] = "Full Payment" Then
        cbBI.Visible = False
    ElseIf rs![BilledIn] = "Not Paid" Then
        cbBI.ListIndex = 0
    Else
        cbBI.ListIndex = 1
    End If
    txtDP.Text = toMoney(toNumber(rs![DownPayment]))
    
    txtEntry(8).Text = rs![Remarks]
    txtLess.Text = toMoney(rs![Less])
    txtDesc.Text = toMoney(rs![Discount])
    txtTA.Text = toMoney(rs![TotalAmount])
    txtAP.Text = toMoney(rs![AmountPaid])
    txtBal.Text = toMoney(rs![Balance])
    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_AR_InvoiceDetails WHERE InvoiceFK=" & PK & " ORDER BY PK ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            With Grid
                If .Rows = 2 And .TextMatrix(1, 11) = "" Then
                    .TextMatrix(1, 1) = RSDetails![ProductCode]
                    .TextMatrix(1, 2) = RSDetails![Description]
                    .TextMatrix(1, 3) = toMoney(RSDetails![SalesPrice(Each)])
                    .TextMatrix(1, 4) = RSDetails![SoldCases]
                    .TextMatrix(1, 5) = RSDetails![SoldBoxes]
                    .TextMatrix(1, 6) = RSDetails![SoldPieces]
                    .TextMatrix(1, 7) = RSDetails![TotalQty]
                    .TextMatrix(1, 8) = RSDetails![Disc]
                    .TextMatrix(1, 9) = toMoney(RSDetails![Amount])
                    .TextMatrix(1, 10) = RSDetails![Free]
                    .TextMatrix(1, 11) = RSDetails![PK]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![ProductCode]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Description]
                    .TextMatrix(.Rows - 1, 3) = toMoney(RSDetails![SalesPrice(Each)])
                    .TextMatrix(.Rows - 1, 4) = RSDetails![SoldCases]
                    .TextMatrix(.Rows - 1, 5) = RSDetails![SoldBoxes]
                    .TextMatrix(.Rows - 1, 6) = RSDetails![SoldPieces]
                    .TextMatrix(.Rows - 1, 7) = RSDetails![TotalQty]
                    .TextMatrix(.Rows - 1, 8) = RSDetails![Disc]
                    .TextMatrix(.Rows - 1, 9) = toMoney(RSDetails![Amount])
                    .TextMatrix(.Rows - 1, 10) = RSDetails![Free]
                    .TextMatrix(.Rows - 1, 11) = RSDetails![PK]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 11
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

    cmdNew.Visible = False
    cmdReset.Visible = False
    cmdGenerate.Visible = False
    dtpDate.Visible = False
    txtDate.Visible = True
    picPurchase.Visible = False
    cmdSave.Visible = False
    btnSold.Visible = False
    txtLess.Locked = True

    'Resize and reposition the controls
    Shape3.Top = 2850
    Label11.Top = 2850
    Line1(1).Visible = False
    Line2(1).Visible = False
    Grid.Top = 3150
    Grid.Height = 2800

    Exit Sub
err:
    'Error if encounter a null value
    If err.Number = 94 Then Resume Next
End Sub

Private Sub InitNSD()
    'For Customer
    With nsdCustomer
        .ClearColumn
        .AddColumn "Customer ID", 1794.89
        .AddColumn "Name", 2264.88
        .AddColumn "Address", 2670.23
        .AddColumn "City/Town", 2190.04
        .AddColumn "Province", 2025.07
        .AddColumn "Zip Code", 1299.96
        .AddColumn "Contact Person", 2174.74
        .AddColumn "Disc.%", 800
        .Connection = CN.ConnectionString
        
        .sqlFields = "CustomerID, Name, Address, CityTown, Province, ZipCode, ContactPerson,Discount, Status, PK"
        .sqlTables = "tbl_AR_Customer"
        .sqlSortOrder = "Name ASC"
        
        .BoundField = "PK"
        .PageBy = 25
        .DisplayCol = 2
        
        .setDropWindowSize 7000, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Customer Records"
        
    End With
    'For Product
    With nsdProduct
        .ClearColumn
        .AddColumn "Product Code", 2064.882
        .AddColumn "Description", 4085.26
        .AddColumn "Unit Cost", 1500
        .AddColumn "Sales Price", 1500
        .AddColumn "Pieces Per Box", 0
        .AddColumn "Pieces Per Case", 0
        
        .Connection = CN.ConnectionString
        
        .sqlFields = "ProductCode,Description,UnitCost,SalesPrice,PiecesPerBox,PiecesPerCase,PK,LoadingFK"
        .sqlTables = "qry_IC_dwnLoadingDetails"
        .sqlSortOrder = "ProductCode ASC"
        .sqlwCondition = "LoadingFK = " & LLFK
        .BoundField = "PK"
        .PageBy = 25
        .DisplayCol = 1
        
        .setDropWindowSize 6800, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Loaded Products From " & LLDate
        
    End With

End Sub

Private Sub txtVan_GotFocus()
    HLText txtVan
End Sub

Private Sub txtCusAdd_GotFocus()
    HLText txtCusAdd
End Sub

Private Sub txtCusCP_GotFocus()
    HLText txtCusCP
End Sub

Private Sub txtSP_GotFocus()
    HLText txtSP
End Sub

Private Sub txtTA_GotFocus()
    HLText txtTA
End Sub

Private Sub txtDP_GotFocus()
    HLText txtDP
End Sub

Private Sub txtAP_GotFocus()
    HLText txtAP
End Sub

Private Sub txtLess_Change()
    txtTA.Text = toMoney(cIAmount - toNumber(txtLess.Text))
End Sub

Private Sub txtLess_Validate(Cancel As Boolean)
    txtLess.Text = toMoney(toNumber(txtLess.Text))
End Sub

