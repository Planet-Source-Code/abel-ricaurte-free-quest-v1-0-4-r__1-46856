VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFinal 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5700
   ForeColor       =   &H00C0C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmFinal.frx":0000
   ScaleHeight     =   280
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCha 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1350
      Index           =   1
      Left            =   150
      ScaleHeight     =   1320
      ScaleWidth      =   1320
      TabIndex        =   8
      Top             =   1320
      Width           =   1350
   End
   Begin MSComctlLib.ImageList imlSph 
      Left            =   4980
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   30
      ImageHeight     =   30
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinal.frx":3A78
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinal.frx":3D7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinal.frx":4083
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picEnd 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3300
      Left            =   1650
      ScaleHeight     =   220
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   260
      TabIndex        =   5
      Top             =   150
      Width           =   3900
      Begin VB.TextBox txtEnd 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3000
         Left            =   150
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "frmFinal.frx":43A5
         Top             =   1650
         Width           =   3600
      End
   End
   Begin VB.Timer tmr 
      Interval        =   100
      Left            =   2640
      Top             =   1920
   End
   Begin VB.PictureBox picSign 
      Appearance      =   0  'Flat
      BackColor       =   &H000C0C0C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   2625
      ScaleHeight     =   450
      ScaleWidth      =   450
      TabIndex        =   3
      Top             =   1890
      Width           =   450
   End
   Begin VB.PictureBox picCha 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1350
      Index           =   2
      Left            =   4200
      ScaleHeight     =   1320
      ScaleWidth      =   1320
      TabIndex        =   0
      Top             =   1320
      Width           =   1350
   End
   Begin VB.Label lblRoll 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fight!"
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
      Height          =   450
      Left            =   2115
      TabIndex        =   7
      Top             =   3660
      Width           =   1500
   End
   Begin VB.Label lblStory 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   150
      LinkTimeout     =   100
      TabIndex        =   4
      Top             =   180
      Width           =   5400
   End
   Begin VB.Image imgHP 
      Appearance      =   0  'Flat
      Height          =   120
      Index           =   1
      Left            =   150
      Picture         =   "frmFinal.frx":43AE
      Stretch         =   -1  'True
      Top             =   2745
      Width           =   1350
   End
   Begin VB.Image imgHP 
      Height          =   120
      Index           =   2
      Left            =   4200
      Picture         =   "frmFinal.frx":46C3
      Stretch         =   -1  'True
      Top             =   2745
      Width           =   1350
   End
   Begin VB.Label lblCha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Opponent"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   4200
      TabIndex        =   2
      Top             =   1020
      Width           =   1350
   End
   Begin VB.Label lblCha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Character"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   150
      TabIndex        =   1
      Top             =   1020
      Width           =   1350
   End
End
Attribute VB_Name = "frmFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iSph As Integer

Private Sub Form_Activate()
  Dim strTmp As String
  
  strTmp = ReadINI("Ending", "Battle", strFree)
  If PStr("Find", strTmp, , "Enabled") > 0 Then txtEnd = PStr("Find", strTmp, "Return", "Introduce"): lblRoll = "Ok" Else txtEnd = Replace(ReadINI("Ending", "Win", strFree), "|", vbCrLf): lblRoll = "The End"
  If PStr("Find", strTmp, , "Picture") > 0 Then picCha(2).Picture = LoadPicture(PStr("Find", strTmp, "Return", "Picture"))
  lblCha(2) = PStr("Find", strTmp, "Return", "Opponent")
  imgHP(1).Width = frmMain.picHP.Width
  picCha(1).Picture = frmMain.picCha
  txtEnd.Top = picEnd.Height / 2
  lblCha(1) = frmMain.lblCha
  frmMain.Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblRoll.ForeColor = vbWhite
End Sub

Private Sub lblRoll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblRoll.ForeColor = vbBlack
End Sub

Private Sub tmr_Timer()
  txtEnd.Top = txtEnd.Top - 1
  If iSph = 3 Then iSph = 1 Else iSph = iSph + 1
  picSign.Picture = imlSph.ListImages(iSph).Picture
End Sub

Private Sub lblRoll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Select Case lblRoll
    Case "Ok"
      picEnd.Visible = False
      lblRoll = "Fight!"
      tmr.Enabled = True
      
    Case "Fight!"
      tmr.Enabled = False
      EndingQuest
      lblRoll = IIf(Me.Tag = "", "Ok", "Continue")
    
    Case "Continue"
      lblStory = ""
      lblRoll = "The End"
      If Me.Tag = "Lost" Then txtEnd = vbCrLf & "You start to feel weak...death has found you!!!" & vbCrLf & "The End" Else txtEnd = Replace(ReadINI("Ending", Me.Tag, strFree), "|", vbCrLf)
      picEnd.Visible = True
      tmr.Enabled = True
      txtEnd.Top = 200
      
    Case "The End"
      If Me.Tag = "Win" Then txtEnd = txtEnd & vbCrLf & "( +30 Exp. Points )": frmMain.lblExp = frmMain.lblExp + 30
      If frmHall.Rating = True Then frmHall.Show
      frmMain.txtMain = vbCrLf & txtEnd
      frmMain.Enabled = True
      Unload Me
    End Select
End Sub

Private Sub EndingQuest()
  Dim iQ As Integer
  If iSph = 3 Then iQ = 5: lblStory = "Mutual Damage!" Else iQ = 30: lblStory = IIf(iSph = 1, "You", lblCha(2)) & " take the damage!"
  If Not iSph = 2 Then If iQ < imgHP(1).Width Then imgHP(1).Width = imgHP(1).Width - iQ Else lblStory = "You Lose!": imgHP(1).Width = 1: Me.Tag = "Lost"
  If Not iSph = 1 Then If iQ < imgHP(2).Width Then imgHP(2).Width = imgHP(2).Width - iQ: imgHP(2).Left = imgHP(2).Left + iQ Else lblStory = "You defeated " & lblCha(2): imgHP(1).Width = 90: Me.Tag = "Win"
  frmMain.picHP.Width = imgHP(1).Width
End Sub
