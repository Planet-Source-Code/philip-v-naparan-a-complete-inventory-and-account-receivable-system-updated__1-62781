VERSION 5.00
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmLoadingPrintOp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading Report"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3810
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLoadingPrintOp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   3810
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnPrev 
      Caption         =   "Preview"
      Default         =   -1  'True
      Height          =   315
      Left            =   975
      TabIndex        =   1
      Top             =   975
      Width           =   1335
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2415
      TabIndex        =   2
      Top             =   975
      Width           =   1335
   End
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -450
      TabIndex        =   3
      Top             =   825
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   53
   End
   Begin ctrlNSDataCombo.NSDataCombo nsdLoading 
      Height          =   315
      Left            =   150
      TabIndex        =   0
      Top             =   375
      Width           =   3540
      _ExtentX        =   6244
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
      Caption         =   "Select Loading To Print"
      Height          =   240
      Index           =   4
      Left            =   150
      TabIndex        =   4
      Top             =   150
      Width           =   2565
   End
End
Attribute VB_Name = "frmLoadingPrintOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnPrev_Click()
    If nsdLoading.BoundText = "" Then
        MsgBox "Please select the last loading date.", vbExclamation
        nsdLoading.SetFocus
    Else
    
        GenerateDSN
        With MAIN.CR
            .Reset: MAIN.InitCrys
             .ReportFileName = App.Path & "\Reports\rptLoading.rpt"
            .Connect = "DSN=" & App.Path & "\rptCN.dsn;PWD=philiprj"
        
            .SelectionFormula = "{qry_IC_Loading.PK} = " & nsdLoading.BoundText
            
            .WindowTitle = "Van Loading Report"
    
            .ParameterFields(0) = "prBussAddr;" & CurrBiz.BUSINESS_ADDRESS & ";True"
            .ParameterFields(1) = "prmBussContact;" & CurrBiz.BUSINESS_CONTACT_INFO & ";True"
            .ParameterFields(2) = "prmTitle;VAN LOADING REPORT;True"
                
            .PageZoom 100
            .Action = 1
        End With
        RemoveDSN
    End If
End Sub

Private Sub Form_Load()
    With nsdLoading
        .ClearColumn
        .AddColumn "Loading No", 2000.126
        .AddColumn "Date", 2200.252
        .AddColumn "Van Name", 4000.252
        .AddColumn "", 0

        .Connection = CN.ConnectionString
        
        .sqlFields = "LoadingNo,Date,VanName,VanFK,PK"
        .sqlTables = "qry_IC_dwnLoadingForPickInv"
        .sqlSortOrder = "PK DESC"
        
        .BoundField = "PK"
        .PageBy = 25
        .DisplayCol = 2
        
        .setDropWindowSize 7000, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Loading Dates"
        
    End With
End Sub
