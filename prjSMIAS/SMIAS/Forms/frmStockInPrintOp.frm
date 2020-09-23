VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStockInPrintOp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock-In Report"
   ClientHeight    =   1800
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
   Icon            =   "frmStockInPrintOp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   3660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnPrev 
      Caption         =   "Preview"
      Default         =   -1  'True
      Height          =   315
      Left            =   750
      TabIndex        =   1
      Top             =   1275
      Width           =   1335
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2190
      TabIndex        =   2
      Top             =   1275
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date Receive"
      Height          =   915
      Left            =   75
      TabIndex        =   4
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
         Format          =   19136515
         CurrentDate     =   38207
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   5
         Top             =   375
         Width           =   1215
      End
   End
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -525
      TabIndex        =   3
      Top             =   1125
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   53
   End
End
Attribute VB_Name = "frmStockInPrintOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnPrev_Click()
    GenerateDSN
    With MAIN.CR
        .Reset: MAIN.InitCrys
         .ReportFileName = App.Path & "\Reports\rptStockIn.rpt"
        .Connect = "DSN=" & App.Path & "\rptCN.dsn;PWD=philiprj"
    
        .SelectionFormula = "{qry_IC_StockReceiveByProd.DateReceive} = cdate(" & Format$(dtpFDate.Value, "yyyy,mm,dd") & ")"
        
        .WindowTitle = "Stock-in Report"

        .ParameterFields(0) = "prBussAddr;" & CurrBiz.BUSINESS_ADDRESS & ";True"
        .ParameterFields(1) = "prmBussContact;" & CurrBiz.BUSINESS_CONTACT_INFO & ";True"
        .ParameterFields(2) = "prmTitle;STOCK-IN REPORT;True"
        .ParameterFields(3) = "prmDateIn;" & Format$(dtpFDate.Value, "MMM-dd-yyyy") & ";True"
            
        .PageZoom 100
        .Action = 1
    End With
    RemoveDSN
End Sub

Private Sub Form_Load()
    dtpFDate.Value = Date
End Sub
