VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSelectZipCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Zip Codes"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   120
   ClientWidth     =   7125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelectZipCode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton sel5 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Refresh"
      Top             =   1635
      Width           =   315
   End
   Begin VB.CommandButton sel4 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Delete"
      Top             =   1320
      Width           =   315
   End
   Begin VB.CommandButton sel3 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Edit"
      Top             =   1005
      Width           =   315
   End
   Begin VB.CommandButton sel2 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "New"
      Top             =   690
      Width           =   315
   End
   Begin VB.CommandButton sel1 
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Find"
      Top             =   375
      Width           =   315
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4440
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select"
      Default         =   -1  'True
      Height          =   315
      Left            =   5760
      TabIndex        =   0
      Top             =   3960
      Width           =   1215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3510
      Left            =   465
      TabIndex        =   2
      Top             =   375
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   6191
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sort"
         Object.Width           =   1413
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "City"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Province"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Zip Code"
         Object.Width           =   1834
      EndProperty
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Select a Zip Code"
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
      TabIndex        =   8
      Top             =   75
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   75
      Top             =   75
      Width           =   6915
   End
End
Attribute VB_Name = "frmSelectZipCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public txtCity          As TextBox
Public txtState         As TextBox
Public txtZipCode       As TextBox

Public rs               As New Recordset
Public OPEN_COMMAND     As Integer '0=For pop-up,1=For managing

Private Sub Command1_Click()
    Call selectCurList
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub




Private Sub Form_Load()
    If OPEN_COMMAND = 1 Then
        Me.Height = 4410
        Command1.Visible = False
        Command2.Visible = False
        
        lblTitle.Caption = Me.Caption
    End If
    '-Start setting up the graphics
    With MAIN
        sel1.Picture = .i16x16.ListImages(9).Picture
        sel2.Picture = .i16x16.ListImages(10).Picture
        sel3.Picture = .i16x16.ListImages(11).Picture
        sel4.Picture = .i16x16.ListImages(12).Picture
        sel5.Picture = .i16x16.ListImages(13).Picture
        
        Set ListView1.SmallIcons = .i16x16
        Set ListView1.Icons = .i16x16
    End With
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM tbl_SM_ZipCodeList ORDER BY CityTown ASC", CN, adOpenStatic, adLockOptimistic
    Call reload_rec
End Sub

Public Sub reload_rec()
    rs.Filter = ""
    rs.Requery
    FillListView ListView1, rs, 4, 2, True, True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmSelectZipCode = Nothing
End Sub


Private Sub ListView1_DblClick()
    If OPEN_COMMAND = 0 Then Call selectCurList
End Sub
Private Sub selectCurList()
    If ListView1.ListItems.Count < 1 Then MsgBox "No record to select.", vbExclamation: Exit Sub
    On Error Resume Next
    txtCity.Text = ListView1.SelectedItem.ListSubItems(1)
    txtState.Text = ListView1.SelectedItem.ListSubItems(2)
    txtZipCode.Text = ListView1.SelectedItem.ListSubItems(3)
    Unload Me
End Sub

Private Sub sel1_Click()
    If ListView1.ListItems.Count < 1 Then MsgBox "No record to search.", vbExclamation: Exit Sub
    With frmFind
        Set .srcListView = ListView1
        .show vbModal
    End With
End Sub

Private Sub sel2_Click()
    With frmSelectZipCodeAdd
        .ADD_STATE = True
        .show vbModal
    End With
End Sub

Private Sub sel3_Click()
    If ListView1.ListItems.Count < 1 Then
        MsgBox "There is no record to edit.", vbInformation
        Exit Sub
    End If
    With frmSelectZipCodeAdd
        .ADD_STATE = False
        .CURR_ZIP = ListView1.SelectedItem.ListSubItems(3)
        .show vbModal
    End With
End Sub

Private Sub sel4_Click()
    On Error GoTo err
    With rs
        '-Check if there is no record
        If .RecordCount < 1 Then MsgBox "No record to delete.", vbExclamation: Exit Sub
        '-Confirm deletion of record
        Dim ANS As Integer
        ANS = MsgBox("Are you sure you want to delete the selected record?", vbCritical + vbYesNo, "Confirm Record Delete")
        Me.MousePointer = vbHourglass
        If ANS = vbYes Then
            If isRecordExist("tbl_SM_ZipCodeList", "ZipCode", ListView1.SelectedItem.ListSubItems(3), True) = False Then
                MsgBox "This zip code is no longer exist in the record. Click ok to reload the records!", vbExclamation, "Unable To Edit"
                Me.MousePointer = vbDefault
                reload_rec
                Exit Sub
            End If
            '-Delete the record
            .AbsolutePosition = CInt(ListView1.SelectedItem)
            .Delete
            reload_rec
            MsgBox "Record has been successfully deleted.", vbInformation, "Confirm"
        End If
        ANS = 0
        Me.MousePointer = vbDefault
    End With
    Exit Sub
err:
        prompt_err err, "frmSelectZipCode", "sel4_Click"
        Me.MousePointer = vbDefault
End Sub

Private Sub sel5_Click()
    Call reload_rec
End Sub
