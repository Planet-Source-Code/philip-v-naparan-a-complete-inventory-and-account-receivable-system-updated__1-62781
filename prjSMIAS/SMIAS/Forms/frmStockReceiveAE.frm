VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStockReceiveAE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   7245
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStockReceiveAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7245
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5625
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "0.00"
      Top             =   975
      Width           =   1500
   End
   Begin VB.TextBox lblUC 
      Height          =   285
      Left            =   5625
      TabIndex        =   22
      Text            =   "0.00"
      Top             =   675
      Width           =   1500
   End
   Begin VB.TextBox lblTQty 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   5625
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "0"
      Top             =   375
      Width           =   1500
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   4
      Left            =   1350
      TabIndex        =   6
      Top             =   3000
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1350
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1125
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   3
      Left            =   1350
      TabIndex        =   5
      Text            =   "0"
      Top             =   2550
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1350
      TabIndex        =   4
      Text            =   "0"
      Top             =   2175
      Width           =   1515
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1350
      TabIndex        =   3
      Text            =   "0"
      Top             =   1800
      Width           =   1515
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   150
      TabIndex        =   7
      Top             =   3600
      Width           =   1680
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5790
      TabIndex        =   9
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   4350
      TabIndex        =   8
      Top             =   3600
      Width           =   1335
   End
   Begin MSDataListLib.DataCombo dcProd 
      Height          =   315
      Left            =   1350
      TabIndex        =   1
      Top             =   750
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   285
      Left            =   1350
      TabIndex        =   0
      Top             =   375
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   24707075
      CurrentDate     =   38207
   End
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   27
      Top             =   3450
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   53
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
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
      Left            =   4050
      TabIndex        =   28
      Top             =   75
      Width           =   3015
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Product"
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
      Left            =   150
      TabIndex        =   26
      Top             =   75
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "(Not Available)"
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   2925
      TabIndex        =   25
      Top             =   2175
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label1 
      Caption         =   "(Not Available)"
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   2925
      TabIndex        =   24
      Top             =   1800
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Unit Cost(Each)"
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   4050
      TabIndex        =   19
      Top             =   675
      Width           =   1515
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Qty"
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   4050
      TabIndex        =   18
      Top             =   375
      Width           =   1515
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Reference"
      Height          =   240
      Index           =   3
      Left            =   75
      TabIndex        =   17
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Description"
      ForeColor       =   &H0000011D&
      Height          =   240
      Index           =   2
      Left            =   75
      TabIndex        =   16
      Top             =   1125
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Pieces"
      Height          =   240
      Index           =   14
      Left            =   75
      TabIndex        =   15
      Top             =   2550
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Boxes"
      Height          =   240
      Index           =   13
      Left            =   75
      TabIndex        =   14
      Top             =   2175
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Cases"
      Height          =   240
      Index           =   12
      Left            =   75
      TabIndex        =   13
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Receive"
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
      Left            =   150
      TabIndex        =   12
      Top             =   1500
      Width           =   3015
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Product Code"
      Height          =   240
      Index           =   1
      Left            =   75
      TabIndex        =   11
      Top             =   750
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Date Receive"
      Height          =   240
      Index           =   0
      Left            =   75
      TabIndex        =   10
      Top             =   375
      Width           =   1215
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   75
      Top             =   1500
      Width           =   3765
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
      Left            =   4050
      TabIndex        =   20
      Top             =   975
      Width           =   1515
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   75
      Top             =   75
      Width           =   3765
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   3975
      Top             =   75
      Width           =   3165
   End
End
Attribute VB_Name = "frmStockReceiveAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit

Dim PCase                   As Long 'Pieces per case
Dim PBox                    As Long 'Pieces per box

Dim old_pieces              As Long 'Old pieces value
Dim old_boxes               As Long 'Old boxes value
Dim old_cases               As Long 'Old cases value

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset

Private Sub DisplayForEditing()
    On Error GoTo err
    
    With rs
        dtpDate.Value = .Fields("DateReceive")
        dcProd.BoundText = .Fields("ProductFK")
        DiplayProdInfo
        lblUC.Text = toMoney(.Fields("UnitCost(Each)"))
        
        old_cases = .Fields("Cases")
        old_boxes = .Fields("Boxes")
        old_pieces = .Fields("Pieces")
        
        txtEntry(1).Text = old_cases
        txtEntry(2).Text = old_boxes
        txtEntry(3).Text = old_pieces
        
        txtEntry(4).Text = .Fields("Reference")
        
        .Update
    End With
    
    Exit Sub
err:
        If err.Number = 94 Then Resume Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetEntry()
    txtEntry(1).Text = "0"
    txtEntry(2).Text = "0"
    txtEntry(3).Text = "0"
    
    lblTQty.Text = "0"
    lblUC.Text = "0.00"
    lblTAmount.Text = "0.00"
End Sub

Private Sub ResetFields()
    clearText Me
    
    
    txtEntry(1).Text = "0"
    txtEntry(2).Text = "0"
    txtEntry(3).Text = "0"
    
    lblTQty.Text = "0"
    lblUC.Text = "0.00"
    lblTAmount.Text = "0.00"
    
    dcProd.BoundText = RightSplitUF(dcProd.Tag)
    DiplayProdInfo
    
    dtpDate.SetFocus
End Sub

Private Sub GeneratePK()
    PK = getIndex("tbl_IC_StockReceive")
End Sub

Private Sub cmdSave_Click()
    If is_empty(dcProd) = True Then Exit Sub
    
    If toNumber(lblTQty.Text) = 0 Then
        MsgBox "Please enter some quantity to receive.", vbExclamation
        txtEntry(3).SetFocus
        Exit Sub
    End If
    
    On Error GoTo err
    
    Dim rsProdUpdate As New Recordset
    
    rsProdUpdate.CursorLocation = adUseClient
    rsProdUpdate.Open "SELECT * FROM tbl_IC_Products WHERE PK =" & dcProd.BoundText, CN, adOpenStatic, adLockOptimistic
    
    CN.BeginTrans
    
    
    If State = adStateAddMode Or State = adStatePopupMode Then
        rs.AddNew
        rs.Fields("PK") = PK
        rs.Fields("DateAdded") = Now
        rs.Fields("AddedByFK") = CurrUser.USER_PK
        
        With rsProdUpdate
            ![Cases] = ![Cases] + toNumber(txtEntry(1).Text)
            ![Boxes] = ![Boxes] + toNumber(txtEntry(2).Text)
            ![Pieces] = ![Pieces] + toNumber(txtEntry(3).Text)
            .Update
        End With
        
    Else
        rs.Fields("DateModified") = Now
        rs.Fields("LastUserFK") = CurrUser.USER_PK
        
        With rsProdUpdate
            'For Cases
            If toNumber(txtEntry(1).Text) <> old_cases Then
                If toNumber(txtEntry(1).Text) > old_cases Then
                    ![Cases] = ![Cases] + (toNumber(txtEntry(1).Text) - old_cases)
                Else
                    ![Cases] = ![Cases] - (old_cases - toNumber(txtEntry(1).Text))
                End If
            End If
            'For Boxes
            If toNumber(txtEntry(2).Text) <> old_boxes Then
                If toNumber(txtEntry(2).Text) > old_boxes Then
                    ![Boxes] = ![Boxes] + (toNumber(txtEntry(2).Text) - old_boxes)
                Else
                    ![Boxes] = ![Boxes] - (old_boxes - toNumber(txtEntry(2).Text))
                End If
            End If
            'For pieces
            If toNumber(txtEntry(3).Text) <> old_pieces Then
                If toNumber(txtEntry(3).Text) > old_pieces Then
                    ![Pieces] = ![Pieces] + (toNumber(txtEntry(3).Text) - old_pieces)
                Else
                    ![Pieces] = ![Pieces] - (old_pieces - toNumber(txtEntry(3).Text))
                End If
            End If
            
            .Update
        End With
        
    End If
    
    With rs
        .Fields("DateReceive") = dtpDate.Value
        .Fields("ProductFK") = dcProd.BoundText
        
        .Fields("Cases") = toNumber(txtEntry(1).Text)
        .Fields("Boxes") = toNumber(txtEntry(2).Text)
        .Fields("Pieces") = toNumber(txtEntry(3).Text)
        
        .Fields("TotalQty") = toNumber(lblTQty.Text)
        .Fields("UnitCost(Each)") = toNumber(lblUC.Text)
        
        .Fields("Reference") = txtEntry(4).Text
        
        .Update
    End With
    
    
    
    CN.CommitTrans
    
    Set rsProdUpdate = Nothing
    
    HaveAction = True
    
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
        prompt_err err, Me.Name, "cmdSave_Click()"
End Sub

Private Sub cmdUsrHistory_Click()
    On Error Resume Next
    Dim tDate1 As String
    Dim tDate2 As String
    Dim tUser1 As String
    Dim tUser2 As String
    
    tDate1 = Format$(rs.Fields("DateAdded"), "MMM-dd-yyyy HH:MM AMPM")
    tDate2 = Format$(rs.Fields("DateModified"), "MMM-dd-yyyy HH:MM AMPM")
    
    tUser1 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & rs.Fields("AddedByFK"), "CompleteName")
    tUser2 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & rs.Fields("LastUserFK"), "CompleteName")
    
    MsgBox "Date Added: " & tDate1 & vbCrLf & _
           "Added By: " & tUser1 & vbCrLf & _
           "" & vbCrLf & _
           "Last Modified: " & tDate2 & vbCrLf & _
           "Modified By: " & tUser2, vbInformation, "Modification History"
           
    tDate1 = vbNullString
    tDate2 = vbNullString
    tUser1 = vbNullString
    tUser2 = vbNullString
End Sub

Private Sub dcProd_Click(Area As Integer)
    On Error Resume Next
    If Area = 2 Then
        If dcProd.BoundText <> "" Then
            If State <> adStateEditMode Then ResetEntry
            DiplayProdInfo
        End If
    End If
End Sub

Private Sub Form_Load()
    dtpDate.Value = Date
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM tbl_IC_StockReceive WHERE PK = " & PK, CN, adOpenStatic, adLockOptimistic
    
    'Bind the data combo
    bind_dc "SELECT * FROM tbl_IC_Products", "ProductCode", dcProd, "PK", True
    'Display the product info
    If toNumber(LeftSplitUF(dcProd.Tag)) > 1 Then DiplayProdInfo
    
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        GeneratePK
    Else
        Caption = "Edit Entry"
        DisplayForEditing
        dcProd.Locked = True
        dcProd.BackColor = &HE6FFFF
        dcProd.ForeColor = &H0&
    End If

End Sub

Private Sub DiplayProdInfo()
    Screen.MousePointer = vbHourglass
    
    Dim rsPI As New Recordset
    
    With rsPI
        .CursorLocation = adUseClient
        
        .Open "SELECT * FROM tbl_IC_Products WHERE PK =" & dcProd.BoundText, CN, adOpenStatic, adLockReadOnly
        
        txtEntry(0).Text = ![Description]
        If State <> adStateEditMode Then lblUC.Text = toMoney(toNumber(![UnitCost]))
        PCase = ![PiecesPerCase]
        PBox = ![PiecesPerBox]
        
    End With
    
    Set rsPI = Nothing
    
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

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or adStateEditMode Then frmStockReceive.RefreshRecords
        MAIN.UpdateInfoMsg
    End If
    
    Set frmStockReceiveAE = Nothing
End Sub

Private Sub lblTQty_GotFocus()
    HLText lblTQty
End Sub

Private Sub lblUC_Change()
    lblTAmount.Text = toMoney(toNumber(lblTQty.Text) * toNumber(lblUC.Text))
End Sub

Private Sub lblUC_GotFocus()
    HLText lblUC
End Sub

Private Sub lblTAmount_GotFocus()
    HLText lblTAmount
End Sub


Private Sub lblUC_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub lblUC_Validate(Cancel As Boolean)
    lblUC.Text = toMoney(toNumber(lblUC.Text))
End Sub


Private Sub txtEntry_Change(Index As Integer)
    If Index > 0 And Index < 4 Then
        lblTQty.Text = (toNumber(txtEntry(1).Text) * PCase) + (toNumber(txtEntry(2).Text) * PBox) + toNumber(txtEntry(3).Text)
        
        lblTAmount.Text = toMoney(toNumber(lblTQty.Text) * toNumber(lblUC.Text))
    End If
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index > 0 And Index < 4 Then
        KeyAscii = isNumber(KeyAscii)
    End If
End Sub

Private Sub txtEntry_Validate(Index As Integer, Cancel As Boolean)
    If Index > 0 And Index < 4 Then
        txtEntry(Index).Text = toNumber(txtEntry(Index).Text)
    End If
End Sub
