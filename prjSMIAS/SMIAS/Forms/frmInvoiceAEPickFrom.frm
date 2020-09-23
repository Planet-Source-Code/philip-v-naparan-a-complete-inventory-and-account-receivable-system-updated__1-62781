VERSION 5.00
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#68.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmInvoiceAEPickFrom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Last Loading Date"
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInvoiceAEPickFrom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4095
   StartUpPosition =   2  'CenterScreen
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -450
      TabIndex        =   4
      Top             =   600
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   1125
      TabIndex        =   2
      Top             =   825
      Width           =   1335
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Default         =   -1  'True
      Height          =   315
      Left            =   2550
      TabIndex        =   1
      Top             =   825
      Width           =   1335
   End
   Begin ctrlNSDataCombo.NSDataCombo nsdLastLoading 
      Height          =   315
      Left            =   1275
      TabIndex        =   0
      Top             =   150
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
      Alignment       =   1  'Right Justify
      Caption         =   "Last Loading"
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
      Index           =   4
      Left            =   0
      TabIndex        =   3
      Top             =   150
      Width           =   1215
   End
End
Attribute VB_Name = "frmInvoiceAEPickFrom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub cmdCancel_Click()
    frmInvoiceAE.CloseMe = True
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If nsdLastLoading.BoundText = "" Then
        MsgBox "Please select the last loading date.", vbExclamation
        nsdLastLoading.SetFocus
    Else
        frmInvoiceAE.LLFK = toNumber(nsdLastLoading.BoundText)
        frmInvoiceAE.LLVFK = toNumber(nsdLastLoading.getSelValueAt(4))
        frmInvoiceAE.LLDate = nsdLastLoading.getSelValueAt(2)
        frmInvoiceAE.txtVan.Text = nsdLastLoading.getSelValueAt(3)
        Unload Me
    End If
End Sub

'nsdProduct.sqlwCondition = "LoadingFK = " & tonumber(nsdLastLoading.BoundText)
Private Sub Form_Load()
    'For Loading
    With nsdLastLoading
        .ClearColumn
        .AddColumn "Loading No", 2000.126
        .AddColumn "Date", 2200.252
        .AddColumn "Van Name", 4000.252
        .AddColumn "", 0

        .Connection = CN.ConnectionString
        
        .sqlFields = "LoadingNo,Date,VanName,VanFK,PK"
        .sqlTables = "qry_IC_dwnLoading"
        .sqlSortOrder = "PK DESC"
        
        .BoundField = "PK"
        .PageBy = 25
        .DisplayCol = 2
        
        .setDropWindowSize 6000, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Loading Dates"
        
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then frmInvoiceAE.CloseMe = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmInvoiceAEPickFrom = Nothing
End Sub
