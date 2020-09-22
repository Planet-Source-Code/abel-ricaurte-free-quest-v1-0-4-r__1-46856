VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRoll 
   BackColor       =   &H80000002&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmRoll.frx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrRoll 
      Enabled         =   0   'False
      Interval        =   55
      Left            =   4920
      Top             =   2790
   End
   Begin MSComctlLib.ImageList imlDice 
      Left            =   120
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   72
      ImageHeight     =   72
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoll.frx":15EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoll.frx":1942
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picDie 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1080
      Index           =   0
      Left            =   2310
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   2
      Top             =   1125
      Width           =   1080
      Begin VB.Label lblDie 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   330
         TabIndex        =   7
         Top             =   200
         Width           =   400
      End
   End
   Begin VB.PictureBox picDie 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1080
      Index           =   1
      Left            =   1080
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   1
      Top             =   1125
      Width           =   1080
      Begin VB.Label lblDie 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   330
         TabIndex        =   6
         Top             =   200
         Width           =   400
      End
   End
   Begin VB.PictureBox picDie 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1080
      Index           =   2
      Left            =   3540
      ScaleHeight     =   1080
      ScaleWidth      =   1080
      TabIndex        =   0
      Top             =   1125
      Width           =   1080
      Begin VB.Label lblDie 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   330
         TabIndex        =   8
         Top             =   200
         Width           =   400
      End
   End
   Begin VB.Label lblAdv 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   975
      TabIndex        =   5
      Top             =   450
      Width           =   3750
   End
   Begin VB.Label lblAct 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "lblAct"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   975
      TabIndex        =   3
      Top             =   2400
      Width           =   3750
   End
   Begin VB.Label lblRoll 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ROLL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   765
      TabIndex        =   4
      Top             =   2835
      Width           =   4020
   End
   Begin VB.Image imgRoll 
      Height          =   360
      Left            =   765
      Picture         =   "frmRoll.frx":1CC2
      Top             =   2835
      Visible         =   0   'False
      Width           =   4020
   End
End
Attribute VB_Name = "frmRoll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private iSta As Integer

Private Sub Form_Load()
  Randomize
  iSta = Rnd * 5
  frmMain.Enabled = False
  lblAct = frmMain.lblSta(iSta) & " Vs. Total Rolled"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblRoll.FontBold = False: imgRoll.Visible = False
End Sub

Private Sub lblRoll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblRoll.FontBold = True: imgRoll.Visible = True
End Sub

Private Sub tmrRoll_Timer()
  Static iRoll As Integer
  If iRoll = 2 Then iRoll = 0
  picDie(0).Picture = imlDice.ListImages(iRoll + 1).Picture
  picDie(1).Picture = picDie(0).Picture
  picDie(2).Picture = picDie(0).Picture
  iRoll = iRoll + 1
End Sub

Private Sub lblRoll_Click()
  Select Case lblRoll
    Case "ROLL"
      lblRoll = "STOP"
      tmrRoll.Enabled = True

    Case "STOP"
      Me.Tag = SumDice
      lblRoll = "CONTINUE"
      tmrRoll.Enabled = False
      If frmMain.lblVal(iSta) - Me.Tag < 0 Then Me.Tag = -6: lblAdv = "You Failed" Else lblAdv = "You Succeed"
    
    Case "CONTINUE"
      frmMain.Update CInt(Me.Tag)
      frmMain.Enabled = True
      lblRoll = "ROLL"
      Unload Me
  End Select
End Sub

Private Function SumDice() As Integer
  Randomize
  lblDie(0) = CInt(Rnd * 9) + 1
  picDie(0).Picture = imlDice.ListImages(1).Picture: SumDice = lblDie(0)
  lblDie(1) = CInt(Rnd * 9) + 1
  picDie(1).Picture = imlDice.ListImages(1).Picture: SumDice = SumDice + lblDie(1)
  lblDie(2) = CInt(Rnd * 9) + 1
  picDie(2).Picture = imlDice.ListImages(1).Picture: SumDice = SumDice + lblDie(2)
End Function
