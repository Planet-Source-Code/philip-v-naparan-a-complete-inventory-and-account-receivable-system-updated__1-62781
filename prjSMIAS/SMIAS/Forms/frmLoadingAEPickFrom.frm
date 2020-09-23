VERSION 5.00
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmLoadingAEPickFrom 
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
   Icon            =   "frmLoadingAEPickFrom.frx":0000
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
Attribute VB_Name = "frmLoadingAEPickFrom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public FOR_VAN_INV As Boolean

Private Sub cmdCancel_Click()
    If FOR_VAN_INV = True Then frmVanInventoryAE.CloseMe = True
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If nsdLastLoading.BoundText = "" Then
        MsgBox "Please select the last loading date.", vbExclamation
        nsdLastLoading.SetFocus
    Else
        If FOR_VAN_INV = True Then
            frmVanInventoryAE.LLFK = toNumber(nsdLastLoading.BoundText)
            frmVanInventoryAE.LLDate = nsdLastLoading.getSelValueAt(2)
            frmVanInventoryAE.txtVan.Text = nsdLastLoading.getSelValueAt(3)
            frmVanInventoryAE.txtVan.Tag = nsdLastLoading.getSelValueAt(4)
            Unload Me
        Else
            frmLoadingAE.LLFK = toNumber(nsdLastLoading.BoundText)
            Unload Me
        End If
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
        If FOR_VAN_INV = True Then
            .sqlTables = "qry_IC_dwnLoading"
        Else
            .sqlTables = "qry_IC_dwnLoadingForPickInv"
        End If
        .sqlSortOrder = "PK DESC"
        
        .BoundField = "PK"
        .PageBy = 25
        .DisplayCol = 2
        
        .setDropWindowSize 7000, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Loading Dates"
        
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If FOR_VAN_INV = True Then If UnloadMode = 0 Then frmVanInventoryAE.CloseMe = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmLoadingAEPickFrom = Nothing
End Sub
