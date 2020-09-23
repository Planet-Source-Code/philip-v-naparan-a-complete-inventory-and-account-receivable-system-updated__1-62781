VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmUserRec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Records"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserRec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   8835
   StartUpPosition =   2  'CenterScreen
   Begin SMIAS.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -75
      TabIndex        =   2
      Top             =   375
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   53
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   635
      ButtonWidth     =   1720
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "i16x16"
      DisabledImageList=   "i16x16"
      HotImageList    =   "i16x16"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Find"
            Object.ToolTipText     =   "F1"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            Object.ToolTipText     =   "F2"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "F3"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit"
            Object.ToolTipText     =   "F4"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            Object.ToolTipText     =   "F5"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Object.ToolTipText     =   "F6"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Close"
            Object.ToolTipText     =   "F7"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   5475
      Top             =   3675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserRec.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserRec.frx":0D9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserRec.frx":17AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserRec.frx":21C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserRec.frx":2BD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserRec.frx":35E4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   4860
      Left            =   75
      TabIndex        =   0
      Top             =   450
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   8573
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "User Name"
         Object.Width           =   3598
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Complete Name"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Admin"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuSC 
      Caption         =   "&File"
      Begin VB.Menu Find 
         Caption         =   "&Find"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu New 
         Caption         =   "&Create New"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Edit 
         Caption         =   "&Edit Selected"
         Shortcut        =   {F3}
      End
      Begin VB.Menu Delete 
         Caption         =   "&Delete Selected"
         Shortcut        =   {F4}
      End
      Begin VB.Menu Refresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu Close 
         Caption         =   "&Close"
         Shortcut        =   {F6}
      End
   End
End
Attribute VB_Name = "frmUserRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim rs As New Recordset

Private Sub Close_Click()
    CommandPass 6
End Sub

Private Sub Delete_Click()
    CommandPass 4
End Sub

Private Sub Edit_Click()
    CommandPass 3
End Sub

Private Sub Find_Click()
    CommandPass 1
End Sub

Private Sub Form_Load()
    'Set the graphics needed
    'Set the graphics for the controls
    With MAIN
        'For listview
        Set lvList.SmallIcons = .i16x16
        Set lvList.Icons = .i16x16
    End With

    RefreshRecords
End Sub

Private Sub RefreshRecords()
    Me.Enabled = False
    If rs.State = adStateOpen Then rs.Close
    rs.Open "SELECT UserID,CompleteName,Admin,PK FROM tbl_SM_Users ORDER BY UserID ASC", CN, adOpenStatic, adLockOptimistic
    FillListView lvList, rs, 3, 2, False, True, "PK"
    Me.Enabled = True
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frmUserRec = Nothing
End Sub



Private Sub lvList_DblClick()
    CommandPass 3
End Sub

Private Sub New_Click()
    CommandPass 2
End Sub

Private Sub Refresh_Click()
    CommandPass 5
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1: CommandPass 1
        Case 3: CommandPass 2
        Case 5: CommandPass 3
        Case 7: CommandPass 4
        Case 9: CommandPass 5
        Case 11: CommandPass 6
    End Select
End Sub

Public Sub CommandPass(ByVal IntCmd As Integer)
On Error GoTo err
    Select Case IntCmd
        'Find
        Case 1
            If lvList.ListItems.Count < 1 Then MsgBox "No record to search.", vbExclamation: Exit Sub
            With frmFind
                Set .srcListView = lvList
                .show vbModal
            End With
        'New
        Case 2
            frmUserRecAE.State = adStateAddMode
            frmUserRecAE.show vbModal
        'Edit
        Case 3
            If lvList.ListItems.Count > 0 Then
                If isRecordExist("tbl_SM_Users", "PK", CLng(lvList.SelectedItem.Tag)) = False Then
                    MsgBox "This record has been removed by other user.Click 'OK' button to refresh the records.", vbExclamation
                    RefreshRecords
                    Exit Sub
                Else
                    With frmUserRecAE
                        .State = adStateEditMode
                        .PK = CLng(lvList.SelectedItem.Tag)
                        .show vbModal
                    End With
                End If
            End If
        'Delete
        Case 4
            If CLng(lvList.SelectedItem.Tag) = CurrUser.USER_PK Then
                MsgBox "You cannot remove your own record because you currently using it.", vbExclamation
                Exit Sub
            Else
                If lvList.ListItems.Count > 0 Then
                    If isRecordExist("tbl_SM_Users", "PK", CLng(lvList.SelectedItem.Tag)) = False Then
                        MsgBox "This record has been removed by other user.Click 'OK' button to refresh the records.", vbExclamation
                        RefreshRecords
                        Exit Sub
                    Else
                        Dim ANS As Integer
                        ANS = MsgBox("Are you sure you want to delete the selected record?" & vbCrLf & vbCrLf & "WARNING: You cannot undo this operation.", vbCritical + vbYesNo, "Confirm Record Delete")
                        Me.MousePointer = vbHourglass
                        If ANS = vbYes Then
                            DelRecwSQL "tbl_SM_Users", "PK", "", True, CLng(lvList.SelectedItem.Tag)
                            RefreshRecords
                            MsgBox "Record has been successfully deleted.", vbInformation, "Confirm"
                        End If
                        ANS = 0
                        Me.MousePointer = vbDefault
                    End If
                Else
                    MsgBox "No record to delete.", vbExclamation
                End If
            End If
        'Reload
        Case 5: RefreshRecords
        'Close
        Case 6: Unload Me
    End Select
    Exit Sub
err:
    If err.Number = -2147467259 Then
        MsgBox "You cannot delete this record because it was used by other records! If you want to delete this record" & vbCrLf & _
               "you will first have to delete or change the records that currenly used this record as shown bellow." & vbCrLf & vbCrLf & _
               err.Description, , "Delete Operation Failed!"
        Me.MousePointer = vbDefault
    End If
End Sub


