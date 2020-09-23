VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCustomerPrintOp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Report"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3660
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustomerPrintOp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   3660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnPrev 
      Caption         =   "Preview"
      Default         =   -1  'True
      Height          =   315
      Left            =   750
      TabIndex        =   6
      Top             =   3150
      Width           =   1335
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2190
      TabIndex        =   7
      Top             =   3150
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Print Option"
      Height          =   1440
      Left            =   75
      TabIndex        =   10
      Top             =   1425
      Width           =   3465
      Begin VB.OptionButton obPrintOp 
         Caption         =   "Print By Van"
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Top             =   600
         Width           =   1365
      End
      Begin VB.OptionButton obPrintOp 
         Caption         =   "Print All"
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo dcVan 
         Height          =   315
         Left            =   525
         TabIndex        =   5
         Top             =   900
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Style           =   2
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Text            =   ""
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Van"
         Height          =   240
         Index           =   11
         Left            =   -750
         TabIndex        =   11
         Top             =   900
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Record To Print"
      Height          =   1290
      Left            =   75
      TabIndex        =   9
      Top             =   75
      Width           =   3465
      Begin VB.CheckBox chkNew 
         Caption         =   "New Customer Only"
         Height          =   240
         Left            =   375
         TabIndex        =   1
         Top             =   525
         Width           =   1740
      End
      Begin VB.OptionButton obRecToPrint 
         Caption         =   "Print Customer w/ Balance"
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   2
         Top             =   825
         Width           =   2565
      End
      Begin VB.OptionButton obRecToPrint 
         Caption         =   "Print Customer Records"
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   0
         Top             =   300
         Value           =   -1  'True
         Width           =   2190
      End
   End
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -525
      TabIndex        =   8
      Top             =   3000
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   53
   End
End
Attribute VB_Name = "frmCustomerPrintOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnPrev_Click()
    Dim strSelFormula As String, strTitle As String
    
    GenerateDSN
    With MAIN.CR
        .Reset: MAIN.InitCrys
        
        If obRecToPrint(0).Value = True Then
             .ReportFileName = App.Path & "\Reports\rptCustomer.rpt"
        Else
             .ReportFileName = App.Path & "\Reports\rptCustomerWB.rpt"
        End If
        .Connect = "DSN=" & App.Path & "\rptCN.dsn;PWD=philiprj"
    
        If obPrintOp(1).Value = True Then
            strSelFormula = "{tbl_AR_Customer.VanFK}=" & dcVan.BoundText
            strTitle = "CUSTOMER LIST (" & dcVan.Text & ")"
        Else
            strTitle = "CUSTOMER LIST"
        End If
        
        If chkNew.Value = 1 Then
            If strSelFormula = "" Then
                strSelFormula = "{tbl_AR_Customer.Status}='New'"
            Else
                strSelFormula = strSelFormula & " AND {tbl_AR_Customer.Status}='New'"
            End If
            
            strTitle = "NEW " & strTitle
        End If
        
        If obRecToPrint(1).Value = True Then
            strSelFormula = Replace(strSelFormula, "tbl_AR_Customer", "qry_AR_CustomerWB")
        End If

        .SelectionFormula = strSelFormula
    
        .WindowTitle = strTitle

        .ParameterFields(0) = "prBussAddr;" & CurrBiz.BUSINESS_ADDRESS & ";True"
        .ParameterFields(1) = "prmBussContact;" & CurrBiz.BUSINESS_CONTACT_INFO & ";True"
        .ParameterFields(2) = "prmTitle;" & strTitle & ";True"
            
        .PageZoom 100
        .Action = 1
    End With
    RemoveDSN
    
    strSelFormula = vbNullString: strTitle = vbNullString
End Sub

Private Sub Form_Load()
    bind_dc "SELECT * FROM tbl_AR_Van", "VanName", dcVan, "PK", True
End Sub

Private Sub obPrintOp_Click(Index As Integer)
    If obPrintOp(1).Value = True Then
        dcVan.Enabled = True
    Else
        dcVan.Enabled = False
    End If
End Sub

Private Sub obRecToPrint_Click(Index As Integer)
    If obRecToPrint(0).Value = True Then
        chkNew.Enabled = True
    Else
        chkNew.Value = 0
        chkNew.Enabled = False
    End If
End Sub
