VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDueChecksPrintOp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Option"
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
   Icon            =   "frmDueChecksPrintOp.frx":0000
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
      TabIndex        =   5
      Top             =   3150
      Width           =   1335
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2190
      TabIndex        =   6
      Top             =   3150
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Print Option"
      Height          =   1440
      Left            =   75
      TabIndex        =   9
      Top             =   1425
      Width           =   3465
      Begin VB.OptionButton obPrintOp 
         Caption         =   "Print By Van"
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   600
         Width           =   1365
      End
      Begin VB.OptionButton obPrintOp 
         Caption         =   "Print All"
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   300
         Value           =   -1  'True
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo dcVan 
         Height          =   315
         Left            =   525
         TabIndex        =   4
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
         TabIndex        =   10
         Top             =   900
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date Range"
      Height          =   1290
      Left            =   75
      TabIndex        =   8
      Top             =   75
      Width           =   3465
      Begin MSComCtl2.DTPicker dtpFDate 
         Height          =   285
         Left            =   675
         TabIndex        =   0
         Top             =   375
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMM-dd-yyyy"
         Format          =   44957699
         CurrentDate     =   38207
      End
      Begin MSComCtl2.DTPicker dtpTDate 
         Height          =   285
         Left            =   675
         TabIndex        =   1
         Top             =   750
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "MMM-dd-yyyy"
         Format          =   44957699
         CurrentDate     =   38207
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   12
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   11
         Top             =   375
         Width           =   1215
      End
   End
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -525
      TabIndex        =   7
      Top             =   3000
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   53
   End
End
Attribute VB_Name = "frmDueChecksPrintOp"
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
        .ReportFileName = App.Path & "\Reports\rptCheckMon.rpt"
        .Connect = "DSN=" & App.Path & "\rptCN.dsn;PWD=philiprj"
    
        If obPrintOp(1).Value = True Then
            strSelFormula = "{tbl_AR_PDCManager.VanFK}=" & dcVan.BoundText
            strTitle = "CHECKS LIST (" & dcVan.Text & ")"
        Else
            strTitle = "CHECKS LIST"
        End If

        If strSelFormula = "" Then
            strSelFormula = "{tbl_AR_PDCManager.DateDue} >= cdate(" & Format$(dtpFDate.Value, "yyyy,mm,dd") & ") AND" & _
                            "{tbl_AR_PDCManager.DateDue} <= cdate(" & Format$(dtpTDate.Value, "yyyy,mm,dd") & ")"
        Else
            strSelFormula = strSelFormula & " AND {tbl_AR_PDCManager.DateDue} >= cdate(" & Format$(dtpFDate.Value, "yyyy,mm,dd") & ") AND" & _
                            "{tbl_AR_PDCManager.DateDue} <= cdate(" & Format$(dtpTDate.Value, "yyyy,mm,dd") & ")"
        End If
        
        .SelectionFormula = strSelFormula
    
        .WindowTitle = "Checks List"

        .ParameterFields(0) = "prBussAddr;" & CurrBiz.BUSINESS_ADDRESS & ";True"
        .ParameterFields(1) = "prmBussContact;" & CurrBiz.BUSINESS_CONTACT_INFO & ";True"
        .ParameterFields(2) = "prmTitle;" & strTitle & ";True"
        .ParameterFields(3) = "prmDateRange;" & Format$(dtpFDate.Value, "MMM-dd-yyyy") & " To " & Format$(dtpTDate.Value, "MMM-dd-yyyy") & ";True"
            
        .PageZoom 100
        .Action = 1
    End With
    RemoveDSN
    
    strSelFormula = vbNullString: strTitle = vbNullString
End Sub

Private Sub Form_Load()
    dtpFDate.Value = Date
    dtpTDate.Value = Date

    bind_dc "SELECT * FROM tbl_AR_Van", "VanName", dcVan, "PK", True
End Sub

Private Sub obPrintOp_Click(Index As Integer)
    If obPrintOp(1).Value = True Then
        dcVan.Enabled = True
    Else
        dcVan.Enabled = False
    End If
End Sub

