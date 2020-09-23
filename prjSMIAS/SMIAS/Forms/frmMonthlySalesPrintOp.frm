VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmMonthlySalesPrintOp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monthly Sales Report"
   ClientHeight    =   2040
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
   Icon            =   "frmMonthlySalesPrintOp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   3660
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbMonth 
      Height          =   315
      Left            =   750
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   150
      Width           =   2790
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   750
      TabIndex        =   1
      Text            =   "0000"
      Top             =   525
      Width           =   765
   End
   Begin VB.CommandButton btnPrev 
      Caption         =   "Preview"
      Default         =   -1  'True
      Height          =   315
      Left            =   750
      TabIndex        =   3
      Top             =   1500
      Width           =   1335
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2190
      TabIndex        =   4
      Top             =   1500
      Width           =   1335
   End
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -525
      TabIndex        =   5
      Top             =   1350
      Width           =   4290
      _ExtentX        =   7567
      _ExtentY        =   53
   End
   Begin MSDataListLib.DataCombo dcVan 
      Height          =   315
      Left            =   750
      TabIndex        =   2
      Top             =   900
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
      Caption         =   "Year"
      Height          =   240
      Index           =   1
      Left            =   -525
      TabIndex        =   8
      Top             =   525
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Month"
      Height          =   240
      Index           =   0
      Left            =   -525
      TabIndex        =   7
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Van"
      Height          =   240
      Index           =   11
      Left            =   -525
      TabIndex        =   6
      Top             =   900
      Width           =   1215
   End
End
Attribute VB_Name = "frmMonthlySalesPrintOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnPrev_Click()

    GenSummaryDailyRpt
    GenerateDSN
    With MAIN.CR
        .Reset: MAIN.InitCrys
        .ReportFileName = App.Path & "\Reports\rptMonthlySales.rpt"

        .Connect = "DSN=" & App.Path & "\rptCN.dsn;PWD=philiprj"

        .WindowTitle = "Monthly Sales Report"

        .ParameterFields(0) = "prBussAddr;" & CurrBiz.BUSINESS_ADDRESS & ";True"
        .ParameterFields(1) = "prmBussContact;" & CurrBiz.BUSINESS_CONTACT_INFO & ";True"
        .ParameterFields(2) = "prmTitle;MONTHLY SALES REPORT;True"
        .ParameterFields(3) = "prVanName;" & dcVan.Text & ";True"
        .ParameterFields(4) = "prmDateRange;" & cbMonth.Text & ";True"

        .PageZoom 100
        .Action = 1
    End With
    RemoveDSN

End Sub

Private Sub GenSummaryDailyRpt()
    Dim rsRpt As New Recordset
    Dim end_day As Byte, c As Byte
    
    rsRpt.CursorLocation = adUseClient
    rsRpt.Open "SELECT * FROM TBL_RPT_DATE ORDER BY Sort ASC", CN, adOpenStatic, adLockOptimistic

    end_day = getEndDay(CDate((cbMonth.ListIndex + 1) & "/" & Format$(Date, "d/yyyy")))
    
    CN.BeginTrans
    rsRpt.MoveFirst
    For c = 1 To 31 'end_day
        If c <= end_day Then
            rsRpt.Fields("Date") = CDate(cbMonth.ListIndex + 1 & "/" & c & "/" & txtYear.Text)
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
    Dim m As Byte
    For m = 1 To 12
        cbMonth.AddItem Format$(m & "/1/1990", "MMMM")
    Next m
    m = 0
    cbMonth.ListIndex = 0

    txtYear.Text = Format$(Date, "YYYY")
    bind_dc "SELECT * FROM tbl_AR_Van", "VanName", dcVan, "PK", True
End Sub

Private Sub txtYear_Change()
    If Val(txtYear.Text) < 2000 Then txtYear.Text = Format$(Date, "YYYY"): txtYear.SelStart = Len(txtYear.Text)
End Sub

Private Sub txtYear_GotFocus()
    HLText txtYear
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

