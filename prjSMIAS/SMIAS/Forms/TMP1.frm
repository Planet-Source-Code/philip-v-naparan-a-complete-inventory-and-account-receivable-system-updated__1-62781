VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form TMP1 
   Caption         =   "Form1"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   ScaleHeight     =   7035
   ScaleWidth      =   9810
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "copy"
      Height          =   690
      Left            =   6300
      TabIndex        =   4
      Top             =   5475
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   465
      Left            =   8250
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1875
      Width           =   1440
   End
   Begin VB.CommandButton Command1 
      Caption         =   "generate"
      Height          =   840
      Left            =   6150
      TabIndex        =   1
      Top             =   4275
      Width           =   1590
   End
   Begin VB.TextBox Text1 
      Height          =   2790
      Left            =   1125
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   3675
      Width           =   4740
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3435
      Left            =   375
      TabIndex        =   3
      Top             =   150
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   6059
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
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
         Text            =   "Loading No"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Van Name"
         Object.Width           =   7056
      EndProperty
   End
End
Attribute VB_Name = "TMP1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
NSD_LV_WIDTH lvList, Text1
End Sub

Private Sub Command2_Click()
Clipboard.Clear
Clipboard.SetText Text1.Text
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
MsgBox KeyAscii
End Sub
