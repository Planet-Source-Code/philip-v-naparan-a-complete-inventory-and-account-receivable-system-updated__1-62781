VERSION 5.00
Begin VB.Form frmStockMonitoringUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Adjustment"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStockMonitoringUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Reset"
      Height          =   315
      Left            =   150
      TabIndex        =   8
      Top             =   3420
      Width           =   1035
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1350
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   525
      Width           =   2925
   End
   Begin VB.TextBox Text1 
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
      Left            =   1350
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   1500
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   4
      Left            =   1395
      TabIndex        =   6
      Text            =   "0"
      Top             =   2475
      Width           =   1515
   End
   Begin VB.TextBox lblTQty 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1395
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0"
      Top             =   2115
      Width           =   1500
   End
   Begin VB.TextBox lblTAmount 
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
      Left            =   1395
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   2805
      Width           =   1500
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   1725
      TabIndex        =   9
      Top             =   3420
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2985
      TabIndex        =   10
      Top             =   3420
      Width           =   1260
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1395
      TabIndex        =   2
      Text            =   "0"
      Top             =   1005
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1395
      TabIndex        =   3
      Text            =   "0"
      Top             =   1380
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   3
      Left            =   1395
      TabIndex        =   4
      Text            =   "0"
      Top             =   1740
      Width           =   1515
   End
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -450
      TabIndex        =   21
      Top             =   3225
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   53
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Description"
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
      Left            =   75
      TabIndex        =   20
      Top             =   525
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Product "
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
      Left            =   75
      TabIndex        =   19
      Top             =   150
      Width           =   1290
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "B.O. in Pieces"
      Height          =   240
      Index           =   3
      Left            =   150
      TabIndex        =   18
      Top             =   2475
      Width           =   1215
   End
   Begin VB.Label Label7 
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
      Left            =   -180
      TabIndex        =   17
      Top             =   2805
      Width           =   1515
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Qty"
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   -105
      TabIndex        =   16
      Top             =   2115
      Width           =   1425
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Cases"
      Height          =   240
      Index           =   12
      Left            =   120
      TabIndex        =   15
      Top             =   1005
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Boxes"
      Height          =   240
      Index           =   13
      Left            =   120
      TabIndex        =   14
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Pieces"
      Height          =   240
      Index           =   14
      Left            =   120
      TabIndex        =   13
      Top             =   1740
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "(Not Available)"
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   2970
      TabIndex        =   12
      Top             =   1005
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label2 
      Caption         =   "(Not Available)"
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   2970
      TabIndex        =   11
      Top             =   1380
      Visible         =   0   'False
      Width           =   1290
   End
End
Attribute VB_Name = "frmStockMonitoringUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public PK As Long

Dim PCase As Long
Dim PBox As Long

Dim qtyCase As Long
Dim qtyBox As Long
Dim qtyPcs As Long
Dim qtyPcsBO As Long

Dim UCost As Double

Dim rsProd As New Recordset
Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    If toNumber(txtEntry(4).Text) > toNumber(lblTQty.Text) Then
        MsgBox "Note: Bad orders must not be more than to your total stock quantity.", vbExclamation
        txtEntry(4).SetFocus
        Exit Sub
    End If
    
    On Error GoTo err
    
    
    With rsProd
        
        .Fields("Cases") = toNumber(txtEntry(1).Text)
        .Fields("Boxes") = toNumber(txtEntry(2).Text)
        .Fields("Pieces") = toNumber(txtEntry(3).Text)
        .Fields("BO") = toNumber(txtEntry(4).Text)
    
        .Update
    End With
    
    
    HaveAction = True
    
    MsgBox "Update in stock record has been successfull.", vbInformation
    Unload Me
    
    Exit Sub
err:
        prompt_err err, Me.Name, "cmdSave_Click()"
End Sub

Private Sub Command1_Click()
    txtEntry(1).Text = qtyCase
    txtEntry(2).Text = qtyBox
    txtEntry(3).Text = qtyPcs
    
    txtEntry(4).Text = qtyPcsBO
End Sub

Private Sub Form_Load()
    DiplayProdInfo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        frmStockMonitoring.RefreshRecords
        MAIN.UpdateInfoMsg
    End If
    
    Set frmStockMonitoringUpdate = Nothing
End Sub


Private Sub DiplayProdInfo()
    Screen.MousePointer = vbHourglass
    
    With rsProd
        .CursorLocation = adUseClient
        
        .Open "SELECT * FROM tbl_IC_Products WHERE PK =" & PK, CN, adOpenStatic, adLockOptimistic
        
        Text1.Text = ![ProductCode]
        Text2.Text = ![Description]
        
        PCase = ![PiecesPerCase]
        PBox = ![PiecesPerBox]
        
        qtyCase = ![Cases]
        qtyBox = ![Boxes]
        qtyPcs = ![Pieces]
        qtyPcsBO = ![BO]
        
        UCost = toNumber(![UnitCost])
        
        Command1_Click
        
    End With
    
    If PCase = 0 Then
        Label1.Visible = True
        txtEntry(1).BackColor = &HE6FFFF
        txtEntry(1).ForeColor = &H0&
        txtEntry(1).Locked = True
    Else
        Label1.Visible = False
        txtEntry(1).BackColor = &H80000005
        txtEntry(1).ForeColor = &H80000008
        txtEntry(1).Locked = False
    End If
    
    If PBox = 0 Then
        Label2.Visible = True
        txtEntry(2).BackColor = &HE6FFFF
        txtEntry(2).ForeColor = &H0&
        txtEntry(2).Locked = True
    Else
        Label2.Visible = False
        txtEntry(2).BackColor = &H80000005
        txtEntry(2).ForeColor = &H80000008
        txtEntry(2).Locked = False
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub txtEntry_Change(Index As Integer)
    If Index > 0 And Index < 5 Then
        lblTQty.Text = (toNumber(txtEntry(1).Text) * PCase) + (toNumber(txtEntry(2).Text) * PBox) + toNumber(txtEntry(3).Text)
        
        lblTAmount.Text = Format$(toNumber(lblTQty.Text) * UCost, "#,##0.00")
    End If
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index > 0 And Index < 5 Then
        KeyAscii = isNumber(KeyAscii)
    End If
End Sub

Private Sub txtEntry_Validate(Index As Integer, Cancel As Boolean)
    If Index > 0 And Index < 5 Then
        txtEntry(Index).Text = toNumber(txtEntry(Index).Text)
    End If
End Sub

Private Sub Text1_GotFocus()
    HLText Text1
End Sub

Private Sub Text2_GotFocus()
    HLText Text2
End Sub

Private Sub lblTAmount_GotFocus()
    HLText lblTAmount
End Sub

Private Sub lblTQty_GotFocus()
    HLText lblTQty
End Sub
