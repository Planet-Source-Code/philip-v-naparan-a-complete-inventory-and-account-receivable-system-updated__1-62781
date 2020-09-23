VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDailySalesPrintOp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daily Sales Report"
   ClientHeight    =   4050
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
   Icon            =   "frmDailySalesPrintOp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   3660
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Print Option"
      Height          =   840
      Left            =   75
      TabIndex        =   12
      Top             =   2400
      Width           =   3465
      Begin MSDataListLib.DataCombo dcVan 
         Height          =   315
         Left            =   525
         TabIndex        =   4
         Top             =   300
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         _Version        =   393216
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
         TabIndex        =   13
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Record To Print"
      Height          =   990
      Left            =   75
      TabIndex        =   11
      Top             =   1350
      Width           =   3465
      Begin VB.OptionButton obRecToPrint 
         Caption         =   "Detailed Of Daily Sales Report"
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   600
         Width           =   3240
      End
      Begin VB.OptionButton obRecToPrint 
         Caption         =   "Summary Of Daily Sales Report"
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   300
         Value           =   -1  'True
         Width           =   2940
      End
   End
   Begin VB.CommandButton btnPrev 
      Caption         =   "Preview"
      Default         =   -1  'True
      Height          =   315
      Left            =   750
      TabIndex        =   5
      Top             =   3525
      Width           =   1335
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2190
      TabIndex        =   6
      Top             =   3525
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Date Range"
      Height          =   1215
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
         Format          =   22675459
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
         Format          =   22675459
         CurrentDate     =   38207
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   10
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   240
         Index           =   1
         Left            =   150
         TabIndex        =   9
         Top             =   375
         Width           =   1215
      End
   End
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -525
      TabIndex        =   7
      Top             =   3375
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   53
   End
End
Attribute VB_Name = "frmDailySalesPrintOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnPrev_Click()
    If DateDiff("d", dtpFDate.Value, dtpTDate.Value) < 0 Then
        MsgBox "Invalid date range.Please check it!", vbExclamation
        Exit Sub
        dtpFDate.SetFocus
    End If

    If obRecToPrint(0).Value = True Then
        If Month(dtpFDate.Value) <> Month(dtpTDate.Value) Then
            MsgBox "Please select a date that have a same month.", vbExclamation
            Exit Sub
            dtpFDate.SetFocus
        End If
    
        If DateDiff("yyyy", dtpTDate.Value, dtpFDate.Value) < 0 Then
            MsgBox "Please select a date that have a same year.", vbExclamation
            Exit Sub
            dtpFDate.SetFocus
        End If
        
        GenSummaryDailyRpt
    End If

    Dim strSelFormula As String, strTitle As String
    
    GenerateDSN
    With MAIN.CR
        .Reset: MAIN.InitCrys
        If obRecToPrint(0).Value = True Then
            .ReportFileName = App.Path & "\Reports\rptDSRDetailed.rpt"
        Else
            .ReportFileName = App.Path & "\Reports\rptDailySales.rpt"
        End If
        
        .Connect = "DSN=" & App.Path & "\rptCN.dsn;PWD=philiprj"
            
        If obRecToPrint(0).Value = True Then
            strTitle = "DAILY SALES REPORT (" & dcVan.Text & ")"
        Else
            strSelFormula = "{qry_AR_Invoice.Date} >= cdate(" & Format$(dtpFDate.Value, "yyyy,mm,dd") & ")" & _
                         " AND {qry_AR_Invoice.Date} <= cdate(" & Format$(dtpTDate.Value, "yyyy,mm,dd") & ")"

            strSelFormula = strSelFormula & " AND {qry_AR_Invoice.VanFK}=" & dcVan.BoundText
            strTitle = "DAILY SALES REPORT (" & dcVan.Text & ")"
            
            .SelectionFormula = strSelFormula
        End If

        
        .WindowTitle = "Daily Sales Report"
        
        .ParameterFields(0) = "prBussAddr;" & CurrBiz.BUSINESS_ADDRESS & ";True"
        .ParameterFields(1) = "prmBussContact;" & CurrBiz.BUSINESS_CONTACT_INFO & ";True"
        .ParameterFields(2) = "prmTitle;" & strTitle & ";True"

        If obRecToPrint(0).Value = True Then
            .ParameterFields(4) = "prmDateRange;" & Format$(dtpFDate.Value, "MMM-dd-yyyy") & " To " & Format$(dtpTDate.Value, "MMM-dd-yyyy") & ";True"
            .ParameterFields(3) = "prVanName;" & dcVan.Text & ";True"
        Else
            .ParameterFields(4) = "prmDateCovered;" & Format$(dtpFDate.Value, "MMM-dd-yyyy") & " To " & Format$(dtpTDate.Value, "MMM-dd-yyyy") & ";True"
            .ParameterFields(3) = "prmVanName;" & dcVan.Text & ";True"
        End If
            
        .PageZoom 100
        .Action = 1
    End With
    RemoveDSN
    
    strSelFormula = vbNullString: strTitle = vbNullString
End Sub

Private Sub GenSummaryDailyRpt()
    Dim rsRpt As New Recordset
    Dim start_day As Byte, end_day As Byte, c As Byte
    
    rsRpt.CursorLocation = adUseClient
    rsRpt.Open "SELECT * FROM TBL_RPT_DATE ORDER BY Sort ASC", CN, adOpenStatic, adLockOptimistic

    start_day = Val(Day(dtpFDate.Value))
    end_day = Val(Day(dtpTDate.Value))
    
    CN.BeginTrans
    rsRpt.MoveFirst
    For c = 1 To 31
        If c >= start_day And c <= end_day Then
            rsRpt.Fields("Date") = CDate(Month(dtpFDate.Value) & "/" & c & "/" & Year(dtpFDate.Value))
            rsRpt.Fields("Disable") = "N"
            rsRpt.Fields("VanFK") = dcVan.BoundText
        Else
            rsRpt.Fields("Disable") = "Y"
        End If
        rsRpt.Update
        rsRpt.MoveNext
    Next c
    CN.CommitTrans

    end_day = 0: c = 0
    Set rsRpt = Nothing
End Sub

Private Sub Form_Load()
    dtpFDate.Value = Date
    dtpTDate.Value = Date
    
    bind_dc "SELECT * FROM tbl_AR_Van", "VanName", dcVan, "PK", True
End Sub
