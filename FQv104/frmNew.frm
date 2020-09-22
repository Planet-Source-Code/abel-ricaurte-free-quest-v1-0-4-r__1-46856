VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmNew 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "New Quest"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCnt 
      Caption         =   "Continue"
      Default         =   -1  'True
      Height          =   285
      Left            =   4065
      TabIndex        =   11
      Top             =   3225
      Width           =   1545
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   285
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3225
      Width           =   900
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      Height          =   330
      Left            =   5280
      TabIndex        =   31
      Top             =   450
      Width           =   330
   End
   Begin VB.FileListBox filList 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   3180
      TabIndex        =   28
      Top             =   1185
      Width           =   2430
   End
   Begin VB.TextBox txtQuest 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3180
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   450
      Width           =   2100
   End
   Begin VB.Frame Frame1 
      Height          =   3510
      Left            =   105
      TabIndex        =   1
      Top             =   15
      Width           =   3000
      Begin MSComDlg.CommonDialog cdlShow 
         Left            =   150
         Top             =   1110
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   330
         Index           =   1
         Left            =   1545
         TabIndex        =   3
         Tag             =   "Alias"
         Text            =   "Alias"
         Top             =   720
         Width           =   1350
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   330
         Index           =   0
         Left            =   1545
         TabIndex        =   2
         Tag             =   "Name"
         Text            =   "Name"
         Top             =   210
         Width           =   1350
      End
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Height          =   330
         Index           =   2
         Left            =   1545
         TabIndex        =   4
         Tag             =   "Skill"
         Text            =   "Skill"
         Top             =   1230
         Width           =   1350
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   240
         Index           =   0
         Left            =   1545
         TabIndex        =   5
         Top             =   1740
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   423
         _Version        =   393216
         Max             =   1
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   240
         Index           =   1
         Left            =   1545
         TabIndex        =   6
         Top             =   1980
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   423
         _Version        =   393216
         Max             =   1
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   240
         Index           =   2
         Left            =   1545
         TabIndex        =   7
         Top             =   2220
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   423
         _Version        =   393216
         Max             =   1
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   240
         Index           =   3
         Left            =   1545
         TabIndex        =   8
         Top             =   2460
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   423
         _Version        =   393216
         Max             =   1
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   240
         Index           =   4
         Left            =   1545
         TabIndex        =   9
         Top             =   2700
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   423
         _Version        =   393216
         Max             =   1
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown 
         Height          =   240
         Index           =   5
         Left            =   1545
         TabIndex        =   10
         Top             =   2940
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   423
         _Version        =   393216
         Max             =   1
         Orientation     =   1
         Enabled         =   -1  'True
      End
      Begin VB.PictureBox picNew 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1350
         Left            =   120
         MouseIcon       =   "frmNew.frx":0000
         MousePointer    =   99  'Custom
         ScaleHeight     =   88
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   88
         TabIndex        =   12
         Top             =   210
         Width           =   1350
      End
      Begin VB.Label lblVal 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   2700
         TabIndex        =   26
         Top             =   1740
         Width           =   90
      End
      Begin VB.Label lblVal 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   2700
         TabIndex        =   25
         Top             =   1980
         Width           =   90
      End
      Begin VB.Label lblVal 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   2700
         TabIndex        =   24
         Top             =   2220
         Width           =   90
      End
      Begin VB.Label lblVal 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   3
         Left            =   2700
         TabIndex        =   23
         Top             =   2460
         Width           =   90
      End
      Begin VB.Label lblVal 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   4
         Left            =   2700
         TabIndex        =   22
         Top             =   2700
         Width           =   90
      End
      Begin VB.Label lblPts 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "75"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2655
         TabIndex        =   21
         Top             =   3210
         Width           =   180
      End
      Begin VB.Label lblVal 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   5
         Left            =   2700
         TabIndex        =   20
         Top             =   2940
         Width           =   90
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Knowledge"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   1740
         Width           =   810
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perception"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   1980
         Width           =   780
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Strengh"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   2220
         Width           =   540
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Toughness"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   2460
         Width           =   795
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agility"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   2700
         Width           =   405
      End
      Begin VB.Label lblStat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dexterity"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   2940
         Width           =   615
      End
      Begin VB.Label Points 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Points Remaining :"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   3210
         Width           =   1350
      End
   End
   Begin VB.Label Character 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Character"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   3180
      TabIndex        =   30
      Top             =   900
      Width           =   2430
   End
   Begin VB.Label Quest 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select Quest"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   3180
      TabIndex        =   29
      Top             =   150
      Width           =   2430
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next
  If Dir$(App.Path & "\Data\") = "" Then MkDir App.Path & "\Data\"
  filList.Path = App.Path & "\Data\"
End Sub

Private Sub picNew_Click()
On Error Resume Next
  cdlShow.Filter = "All Picture Files|*.bmp;*.dib;*.gif;*.jpg;*.wmf;*.emf;*.ico"
  cdlShow.InitDir = App.Path
  cdlShow.filename = ""
  cdlShow.ShowOpen
  If Err = 0 Then picNew.Tag = cdlShow.filename: picNew.Picture = LoadPicture(picNew.Tag) Else picNew.Tag = ""
End Sub

Private Sub txtData_GotFocus(Index As Integer)
  If txtData(Index) = txtData(Index).Tag Then txtData(Index) = ""
End Sub

Private Sub txtData_LostFocus(Index As Integer)
  If txtData(Index) = "" Then txtData(Index) = txtData(Index).Tag
End Sub

Private Sub UpDown_DownClick(Index As Integer)
  If lblVal(Index) > 0 And lblPts < 100 Then lblVal(Index) = lblVal(Index) - 5: lblPts = lblPts + 5
End Sub

Private Sub UpDown_UpClick(Index As Integer)
  If lblVal(Index) < 25 And lblPts > 0 Then lblVal(Index) = lblVal(Index) + 5: lblPts = lblPts - 5
End Sub

Private Sub cmdOpen_Click()
On Error Resume Next
  cdlShow.Filter = "Free-Quest File (*.qst)|*.qst"
  cdlShow.InitDir = App.Path
  cdlShow.filename = ""
  cdlShow.ShowOpen
  If Err = 0 Then txtQuest = cdlShow.filename
End Sub

Private Sub filList_Click()
On Error Resume Next
  txtData(0) = filList.filename
  lblPts = 75
  
  varTmp = Split(ReadINI("Character", "Data", filList.Path & "\" & txtData(0)), "")
  If UBound(varTmp) = 3 Then txtData(1) = varTmp(1): txtData(2) = varTmp(2): picNew.Tag = varTmp(3) Else picNew.Tag = ""
  picNew.Picture = LoadPicture(picNew.Tag)
  
  varTmp = Split(ReadINI("Character", "Stats", filList.Path & "\" & txtData(0)), "|")
  If UBound(varTmp) = -1 Then varTmp = Split("|||||", "|")
  For i% = 0 To 5
    lblVal(i%) = varTmp(i%): lblPts = lblPts - varTmp(i%)
  Next i%
End Sub

Private Sub filList_KeyDown(KeyCode As Integer, Shift As Integer)
  Dim iMsg As Integer
  If filList.ListIndex = -1 Or KeyCode <> vbKeyDelete Then Exit Sub
  iMsg = MsgBox("Delete '" & filList.List(filList.ListIndex) & "'?", vbOKCancel + vbQuestion)
  If iMsg = 1 Then Kill filList.Path & "\" & filList.filename: filList.Refresh: cmdNew_Click
End Sub

Private Sub cmdNew_Click()
  For i% = 0 To 5
    If i% < 3 Then txtData(i%) = txtData(i%).Tag
    lblVal(i%) = 0
  Next i%
  
  picNew.Picture = LoadPicture()
  txtData(0).SetFocus
  lblPts = 75
End Sub

Private Sub cmdCnt_Click()
  If txtQuest = "" Then MsgBox "You must select a Quest.": Exit Sub
  If txtData(0) = "Name" Then MsgBox "You must give the Character a name.": txtData(0).SetFocus: Exit Sub
  
  For i% = -1 To 5
    If i% = -1 Then Me.Tag = "" Else Me.Tag = Me.Tag & lblVal(i%) & "|"
  Next i%
  
  WriteINI "Character", "Data", txtData(0) & "" & txtData(1) & "" & txtData(2) & "" & picNew.Tag, filList.Path & "\" & txtData(0)
  WriteINI "Character", "Stats", Me.Tag & "0", filList.Path & "\" & txtData(0)
  frmMain.Starting txtQuest, txtData(0)
  filList.Refresh
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmMain.Show
  Cancel = 1
  Me.Hide
End Sub
