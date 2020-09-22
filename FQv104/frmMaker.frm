VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMaker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: Quest-Maker ::"
   ClientHeight    =   6735
   ClientLeft      =   150
   ClientTop       =   555
   ClientWidth     =   8700
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMaker.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   449
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   580
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstAct 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      ItemData        =   "frmMaker.frx":08CA
      Left            =   2400
      List            =   "frmMaker.frx":08CC
      TabIndex        =   28
      Top             =   3360
      Visible         =   0   'False
      Width           =   2340
   End
   Begin MSComDlg.CommonDialog cdlShow 
      Left            =   8235
      Top             =   6270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000007&
      Height          =   1500
      Left            =   5700
      Picture         =   "frmMaker.frx":08CE
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   27
      Top             =   5220
      Width           =   1500
      Begin VB.Image imgDot 
         Height          =   225
         Left            =   630
         Picture         =   "frmMaker.frx":19FB
         Top             =   645
         Width           =   225
      End
      Begin VB.Image imgDown 
         Height          =   180
         Left            =   630
         Picture         =   "frmMaker.frx":1ADF
         Top             =   675
         Width           =   225
      End
      Begin VB.Image img 
         Height          =   510
         Index           =   0
         Left            =   15
         Top             =   975
         Width           =   510
      End
      Begin VB.Image img 
         Height          =   510
         Index           =   1
         Left            =   495
         Top             =   975
         Width           =   510
      End
      Begin VB.Image img 
         Height          =   510
         Index           =   2
         Left            =   975
         Top             =   975
         Width           =   510
      End
      Begin VB.Image img 
         Height          =   510
         Index           =   3
         Left            =   15
         Top             =   495
         Width           =   510
      End
      Begin VB.Image img 
         Height          =   510
         Index           =   4
         Left            =   495
         Top             =   495
         Width           =   510
      End
      Begin VB.Image img 
         Height          =   510
         Index           =   5
         Left            =   975
         Top             =   495
         Width           =   510
      End
      Begin VB.Image img 
         Height          =   510
         Index           =   6
         Left            =   15
         Top             =   15
         Width           =   510
      End
      Begin VB.Image img 
         Height          =   510
         Index           =   7
         Left            =   495
         Top             =   15
         Width           =   510
      End
      Begin VB.Image img 
         Height          =   510
         Index           =   8
         Left            =   975
         Top             =   15
         Width           =   510
      End
   End
   Begin VB.TextBox txtRst 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   2250
      TabIndex        =   25
      Top             =   5220
      Width           =   3375
   End
   Begin VB.TextBox txtRst 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   2250
      TabIndex        =   24
      Top             =   6390
      Width           =   3375
   End
   Begin VB.TextBox txtRst 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   2250
      TabIndex        =   23
      Top             =   6000
      Width           =   3375
   End
   Begin VB.ComboBox cmbRst 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      ItemData        =   "frmMaker.frx":1BAE
      Left            =   570
      List            =   "frmMaker.frx":1BB5
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   5220
      Width           =   1650
   End
   Begin VB.ComboBox cmbRst 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      ItemData        =   "frmMaker.frx":1BC1
      Left            =   570
      List            =   "frmMaker.frx":1BC8
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   6390
      Width           =   1650
   End
   Begin VB.ComboBox cmbRst 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      ItemData        =   "frmMaker.frx":1BD4
      Left            =   570
      List            =   "frmMaker.frx":1BDB
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   6000
      Width           =   1650
   End
   Begin VB.TextBox txtRst 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   2250
      TabIndex        =   11
      Top             =   5610
      Width           =   3375
   End
   Begin VB.ComboBox cmbRst 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      ItemData        =   "frmMaker.frx":1BE7
      Left            =   570
      List            =   "frmMaker.frx":1BEE
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   5610
      Width           =   1650
   End
   Begin VB.TextBox txtTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5700
      TabIndex        =   4
      Top             =   4950
      Width           =   1500
   End
   Begin VB.ListBox lstTmp 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   2
      ItemData        =   "frmMaker.frx":1BFA
      Left            =   4740
      List            =   "frmMaker.frx":1C01
      TabIndex        =   3
      Top             =   4230
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.ListBox lstTmp 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   1
      ItemData        =   "frmMaker.frx":1C0D
      Left            =   2400
      List            =   "frmMaker.frx":1C14
      TabIndex        =   2
      Top             =   4230
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.ListBox lstTmp 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Index           =   0
      ItemData        =   "frmMaker.frx":1C20
      Left            =   60
      List            =   "frmMaker.frx":1C27
      TabIndex        =   1
      Top             =   4230
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.TextBox txtMain 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2430
      Index           =   1
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   7
      Tag             =   "Intro"
      Top             =   2475
      Width           =   7200
   End
   Begin VB.Frame Frame1 
      Height          =   4995
      Left            =   7260
      TabIndex        =   6
      Top             =   -75
      Width           =   1440
      Begin VB.CommandButton cmdElm 
         Caption         =   "Other"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   15
         TabIndex        =   10
         Top             =   4680
         Width           =   1405
      End
      Begin VB.CommandButton cmdElm 
         Caption         =   "Character"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   15
         TabIndex        =   9
         Top             =   4395
         Width           =   1405
      End
      Begin VB.CommandButton cmdElm 
         Caption         =   "Object"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   15
         TabIndex        =   8
         Top             =   120
         Width           =   1405
      End
      Begin VB.PictureBox picBack 
         BorderStyle     =   0  'None
         Height          =   3930
         Left            =   30
         ScaleHeight     =   262
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   92
         TabIndex        =   12
         Top             =   420
         Width           =   1380
         Begin VB.ListBox lstSel 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   3150
            Left            =   0
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   420
            Width           =   1380
         End
         Begin VB.Image imgTab 
            Height          =   315
            Index           =   3
            Left            =   1020
            Picture         =   "frmMaker.frx":1C33
            ToolTipText     =   "Add Task"
            Top             =   3585
            Width           =   315
         End
         Begin VB.Image imgTab 
            Height          =   315
            Index           =   2
            Left            =   1020
            Picture         =   "frmMaker.frx":1D31
            ToolTipText     =   "Delete"
            Top             =   45
            Width           =   315
         End
         Begin VB.Image imgTab 
            Height          =   315
            Index           =   1
            Left            =   525
            Picture         =   "frmMaker.frx":1DE2
            ToolTipText     =   "Edit"
            Top             =   45
            Width           =   315
         End
         Begin VB.Image imgTab 
            Height          =   315
            Index           =   0
            Left            =   30
            Picture         =   "frmMaker.frx":1EB9
            ToolTipText     =   "Add"
            Top             =   45
            Width           =   315
         End
         Begin VB.Label Label6 
            Caption         =   "Add Task"
            Height          =   210
            Left            =   75
            TabIndex        =   29
            Top             =   3660
            Width           =   900
         End
      End
   End
   Begin VB.TextBox txtMain 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2430
      Index           =   0
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   0
      Tag             =   "Description"
      Top             =   30
      Width           =   7200
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2280
      TabIndex        =   30
      Top             =   4965
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   315
      Left            =   5280
      Picture         =   "frmMaker.frx":1F68
      Top             =   4905
      Width           =   315
   End
   Begin VB.Image imgCase 
      Height          =   450
      Left            =   7770
      Picture         =   "frmMaker.frx":2054
      Top             =   5730
      Width           =   450
   End
   Begin VB.Image imgAdd 
      Height          =   450
      Left            =   7770
      Picture         =   "frmMaker.frx":25CE
      Top             =   5730
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label lblMove 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add/Move"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7320
      TabIndex        =   26
      Top             =   4950
      Width           =   1350
   End
   Begin VB.Image imgRem 
      Height          =   450
      Left            =   7770
      Picture         =   "frmMaker.frx":2B48
      Top             =   5730
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgArrow 
      Height          =   450
      Index           =   1
      Left            =   7770
      Picture         =   "frmMaker.frx":3134
      Tag             =   "South Room"
      Top             =   6180
      Width           =   450
   End
   Begin VB.Image imgArrow 
      Height          =   450
      Index           =   3
      Left            =   7320
      Picture         =   "frmMaker.frx":31F5
      Tag             =   "West Room"
      Top             =   5730
      Width           =   450
   End
   Begin VB.Image imgArrow 
      Height          =   450
      Index           =   5
      Left            =   8220
      Picture         =   "frmMaker.frx":32AD
      Tag             =   "East Room"
      Top             =   5730
      Width           =   450
   End
   Begin VB.Image imgArrow 
      Height          =   450
      Index           =   7
      Left            =   7770
      Picture         =   "frmMaker.frx":3365
      Tag             =   "North Room"
      Top             =   5280
      Width           =   450
   End
   Begin VB.Image imgRing 
      Height          =   1350
      Left            =   7320
      Picture         =   "frmMaker.frx":3424
      Top             =   5280
      Width           =   1350
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7260
      TabIndex        =   22
      Top             =   4920
      Width           =   1380
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "East"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   75
      TabIndex        =   21
      Top             =   5610
      Width           =   450
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "North"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   75
      TabIndex        =   20
      Top             =   5220
      Width           =   450
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "West"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   75
      TabIndex        =   19
      Top             =   6390
      Width           =   450
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "South"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   75
      TabIndex        =   15
      Top             =   6000
      Width           =   450
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Restrictions"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   600
      TabIndex        =   5
      Top             =   4965
      Width           =   900
   End
   Begin VB.Menu mnuFil 
      Caption         =   "&File"
      Begin VB.Menu mnuFilNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuFilOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuFilBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilSav 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFilSas 
         Caption         =   "Sa&ve As..."
      End
      Begin VB.Menu mnuFilBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilExt 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuQue 
      Caption         =   "&Quest"
      Begin VB.Menu mnuQueOpt 
         Caption         =   "S&tarting..."
         Index           =   4
      End
      Begin VB.Menu mnuQueOpt 
         Caption         =   "&Battle..."
         Index           =   5
      End
      Begin VB.Menu mnuQueOpt 
         Caption         =   "E&nding..."
         Index           =   6
      End
   End
   Begin VB.Menu mnuRet 
      Caption         =   "&Return"
   End
End
Attribute VB_Name = "frmMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private aP() As Variant, iElm As Integer, Pos$

Private Sub Form_Load()
  strFile = App.Path & "\Maker.tmp"
  frmMain.Hide
  mnuFilNew_Click
End Sub

Private Sub mnuFilNew_Click()
On Error Resume Next
  Kill strFile
  cdlShow.filename = ""
  LoadQuest 0, 0
  ShowMap True
End Sub

Private Sub mnuFilOpen_Click()
On Error Resume Next
  cdlShow.Filter = "Free-Quest File (*.qst)|*.qst"
  cdlShow.InitDir = App.Path
  cdlShow.ShowOpen
  If Err = 0 Then FileCopy cdlShow.filename, strFile: LoadQuest 0, 0: ShowMap True
End Sub

Private Sub mnuFilSav_Click()
  If cdlShow.filename = "" Then mnuFilSas_Click Else SaveQuest: FileCopy strFile, cdlShow.filename
End Sub

Private Sub mnuFilSas_Click()
On Error Resume Next
  cdlShow.Filter = "Free-Quest File (*.qst)|*.qst"
  cdlShow.Flags = cdlOFNOverwritePrompt
  cdlShow.InitDir = App.Path
  cdlShow.ShowSave
  If Err = 0 Then SaveQuest: FileCopy strFile, cdlShow.filename: Me.Caption = ":: Quest-Maker :: " & cdlShow.FileTitle
End Sub

Private Sub mnuFilExt_Click()
  End
End Sub

Private Sub mnuQueOpt_Click(Index As Integer)
  SaveQuest
  frmAdd.LoadTab Index, ""
End Sub

Private Sub mnuRet_Click()
  Unload Me
End Sub

Private Sub txtMain_LostFocus(Index As Integer)
  If Trim(Replace(txtMain(Index), vbCrLf, "")) = "" Then txtMain(Index) = txtMain(Index).Tag
End Sub

Private Sub LoadQuest(incX As Integer, incY As Integer)
  Static iX As Integer, iY As Integer, strTmp As String
  Me.Caption = ":: Quest-Maker :: " & IIf(cdlShow.filename = "", "Untitled", cdlShow.FileTitle)
  For i% = 0 To 2
    lstTmp(i%).Clear
  Next i%
  
  strTmp = ReadINI("Starting", "Start", strFile)
  If incX + incY = 0 Then If strTmp = "" Then iX = 0: iY = 0 Else iX = Split(strTmp, ",")(0): iY = Split(strTmp, ",")(1)
  iX = iX + incX: iY = iY + incY: aP = Array(iX, iY): Pos$ = iX & "," & iY
  strTmp = PStr("Find", ReadINI("Map", "Title", strFile), "Return", Pos$)
  txtTitle = IIf(strTmp = "", "Unknown", strTmp)
  strTmp = Replace(ReadINI("Description", Pos$, strFile), "|", vbCrLf)
  txtMain(0) = IIf(strTmp = "", txtMain(0).Tag, strTmp)
  strTmp = Replace(ReadINI("Intro", Pos$, strFile), "|", vbCrLf)
  txtMain(1) = IIf(strTmp = "", txtMain(1).Tag, strTmp)
  varTmp = Split(ReadINI("Element", Pos$, strFile), "|")
  If UBound(varTmp) = 3 Then For e% = 0 To 2: SplitAndAdd lstTmp(e%), varTmp(e%): Next e% 'OJO e% no i%
  Restrictions
  ReloadList lstSel, iElm: ObjToObj lstSel, lstTmp(iElm)
End Sub

Private Sub SaveQuest()
  Dim strTmp As String
  
  For i% = 0 To 2
    For e% = 0 To lstTmp(i%).ListCount - 1
      strTmp = strTmp & "" & lstTmp(i%).List(e%)
    Next e%
    strTmp = strTmp & "|"
  Next i%
  WriteINI "Element", Pos$, strTmp, strFile
    
  If PStr("Find", ReadINI("Map", "Title", strFile), , ReadINI("Starting", "Start", strFile)) = 0 Then WriteINI "Starting", "Start", Pos$, strFile
  WriteINI "Map", "Title", PStr("Add", ReadINI("Map", "Title", strFile), txtTitle, Pos$), strFile
  strTmp = IIf(txtMain(0) = txtMain(0).Tag, "", Replace(txtMain(0), vbCrLf, "|"))
  WriteINI "Description", Pos$, strTmp, strFile
  strTmp = IIf(txtMain(1) = txtMain(1).Tag, "", Replace(txtMain(1), vbCrLf, "|"))
  WriteINI "Intro", Pos$, strTmp, strFile
  
  strTmp = ReadINI("Restriction", "Room", strFile)
  For i% = 0 To 3
    strTmp = PStr("Add", strTmp, IIf(cmbRst(i%).ListIndex > 0, cmbRst(i%).Text, ""), Pos$ & "," & i%)
    WriteINI "Restriction", Pos$ & "," & i%, txtRst(i%), strFile
  Next i%
  WriteINI "Restriction", "Room", strTmp, strFile
  
  strTmp = ReadINI("Element", "Actions", strFile)
  For i% = 0 To lstAct.ListCount - 1
    strTmp = PStr("Add", strTmp, , lstAct.List(i%))
  Next i%
  WriteINI "Element", "Actions", strTmp, strFile
End Sub

Private Sub ShowMap(bCase As Boolean)
  Dim iNum As Integer, iTile As Integer, strTmp As String
  strTmp = ReadINI("Map", "Title", strFile)
  PStr "Add", strTmp, txtTitle, Pos$

  For i% = aP(1) - 1 To aP(1) + 1
    For e% = aP(0) - 1 To aP(0) + 1
      If InStr(strTmp, "|" & e% & "," & i% & "") = 0 Then img(iNum).Visible = False Else img(iNum).Visible = True
      If e% = aP(0) Xor i% = aP(1) Then If img(iNum).Visible = False Then imgArrow(iNum).Visible = bCase Else imgArrow(iNum).Visible = True
      If InStr(strTmp, "|" & e% - 1 & "," & i% & "") > 0 Then iTile = 1 Else iTile = 0
      If InStr(strTmp, "|" & e% & "," & i% + 1 & "") > 0 Then iTile = iTile + 2
      If InStr(strTmp, "|" & e% + 1 & "," & i% & "") > 0 Then iTile = iTile + 4
      If InStr(strTmp, "|" & e% & "," & i% - 1 & "") > 0 Then iTile = iTile + 8
      img(iNum).Picture = frmMain.imlTiles.ListImages(iTile + 1).Picture
      iNum = iNum + 1
    Next e%
  Next i%
End Sub

Private Sub imgArrow_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim iMsg As Integer, iX As Integer, iY As Integer, strTmp As String
  SaveQuest
  If Index = 7 Then iY = 1
  If Index = 5 Then iX = 1
  If Index = 3 Then iX = -1
  If Index = 1 Then iY = -1
  imgArrow(Index).Visible = False
  If lblMove = "Remove" Then iMsg = MsgBox("Delete " & imgArrow(Index).Tag & "?", vbYesNo + vbQuestion, "Quest-Maker") Else LoadQuest iX, iY
  If iMsg = vbYes Then WriteINI "Map", "Title", PStr("Remove", ReadINI("Map", "Title", strFile), , aP(0) + iX & "," & aP(1) + iY), strFile
  imgCase.Picture = imgAdd.Picture
  imgDot.Visible = False
  lblMove = "Add/Move"
  ShowMap True
End Sub

Private Sub imgArrow_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgDot.Visible = True
  ShowMap True
End Sub

Private Sub imgTab_Click(Index As Integer)
  Dim iMsg As Integer, strPhr As String
  If Index > 0 Then If lstSel.ListIndex = -1 Then Exit Sub Else SaveQuest
  
  Select Case Index
    Case 0
      frmAdd.LoadTab iElm, ""
      
    Case 1
      GetIndex lstTmp(iElm), lstSel.Text, "Remove": WriteINI "Element", iElm, PStr("Remove", ReadINI("Element", iElm, strFile), , LCase(lstSel.Text)), strFile
      frmAdd.LoadTab IIf(UBound(Split(lstSel.Text, "/")) = 0, iElm, 3), lstSel.Text
      frmAdd.cmdAdd.Caption = "Update"
      
    Case 2
      iMsg = MsgBox("Delete the " & lstSel.Text & "?", vbYesNo + vbQuestion, "Quest-Maker")
      If iMsg = vbNo Then Exit Sub
      GetIndex lstTmp(iElm), lstSel.Text, "Remove": WriteINI "Element", iElm, PStr("Remove", ReadINI("Element", iElm, strFile), , LCase(lstSel.Text)), strFile
      If UBound(Split(lstSel.Text, "/")) > 0 Then WriteINI "Actions", CStr(Split(lstSel, "/")(0)), PStr("Remove", ReadINI("Actions", CStr(Split(lstSel, "/")(0)), strFile), , CStr(Split(lstSel, "/")(1))), strFile
      ReloadList lstSel, iElm: ObjToObj lstSel, lstTmp(iElm)
    
    Case 3
      frmAdd.txtSub = Split(lstSel.Text, "/")(0)
      frmAdd.LoadTab 3, ""
  End Select
End Sub

Private Sub lstSel_DblClick()
  imgTab_Click (1)
End Sub

Private Sub imgCase_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgCase.Picture = Me.Picture
End Sub

Private Sub imgCase_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If lblMove = "Add/Move" Then lblMove = "Remove": ShowMap False: imgCase.Picture = imgRem.Picture Else lblMove = "Add/Move": ShowMap True: imgCase.Picture = imgAdd.Picture
End Sub

Private Sub cmdElm_Click(Index As Integer)
  If Index < 1 Then cmdElm(1).Top = 4395 Else cmdElm(1).Top = 405
  If Index < 2 Then cmdElm(2).Top = 4680 Else cmdElm(2).Top = 690
  
  picBack.Top = (Index * 300) + 420
  picBack.SetFocus
  iElm = Index
  ReloadList lstSel, iElm: ObjToObj lstSel, lstTmp(iElm)
End Sub

Private Sub lstSel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgTab_MouseMove 5, Button, Shift, X, Y
End Sub

Private Sub txtMain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgTab_MouseMove 5, Button, Shift, X, Y
End Sub

Private Sub imgTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  For i% = 0 To 3
    If i% = Index Then imgTab(i%).BorderStyle = 1 Else imgTab(i%).BorderStyle = 0
  Next i%
End Sub

Private Sub Restrictions()
  For a% = 0 To 3
    ReloadList cmbRst(a%), -1, "<None>"
    cmbRst(a%).AddItem "<No entry>"
    cmbRst(a%).ListIndex = GetIndex(cmbRst(a%), PStr("Find", ReadINI("Restriction", "Room", strFile), "Return", Pos$ & "," & a%))
    txtRst(a%) = ReadINI("Restriction", Pos$ & "," & a%, strFile)
  Next a%
End Sub

Public Sub ReloadList(objTmp As Object, Optional iBox As Integer, Optional strAdd As String)
  objTmp.Clear
  If strAdd <> "" Then objTmp.AddItem strAdd
  For e% = 0 To 2
    For i% = 0 To lstTmp(e%).ListCount - 1
      varTmp = Split(ReadINI("Actions", lstTmp(e%).List(i%), strFile), "|")
      If iBox = -1 Or iBox = e% Then For a% = 1 To UBound(varTmp): objTmp.AddItem lstTmp(e%).List(i%) & "/" & Split(varTmp(a%), "")(0): Next a%
    Next i%
  Next e%
End Sub

Public Sub AddElement(strAdd As String, iTab As Integer)
  If iTab = 3 Then If frmAdd.txtAct <> "" Then GetIndex lstAct, frmAdd.txtAct, "Remove": lstAct.AddItem frmAdd.txtAct
  If iTab < 3 Then If strAdd <> "" Then lstTmp(iElm).AddItem strAdd: WriteINI "Element", iElm, PStr("Add", ReadINI("Element", iElm, strFile), , LCase(strAdd)), strFile
  Restrictions
  ReloadList lstSel, iElm: ObjToObj lstSel, lstTmp(iElm)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmMain.Show
  Cancel = 1
  Hide
End Sub
