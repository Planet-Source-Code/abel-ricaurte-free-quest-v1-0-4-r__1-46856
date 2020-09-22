VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   ":: Free-Quest ::"
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
   Icon            =   "frmMain.frx":0000
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
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      IntegralHeight  =   0   'False
      ItemData        =   "frmMain.frx":08CA
      Left            =   2430
      List            =   "frmMain.frx":08FE
      Sorted          =   -1  'True
      TabIndex        =   33
      Top             =   3360
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   225
      Left            =   75
      TabIndex        =   0
      Top             =   4950
      Width           =   4950
   End
   Begin MSComctlLib.ImageList imlTiles 
      Left            =   8175
      Top             =   6225
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   34
      ImageHeight     =   34
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":096D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C36
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0EFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":148A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":174E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2276
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":253D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":280F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DA9
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":306F
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3344
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrAni 
      Interval        =   750
      Left            =   7260
      Top             =   6300
   End
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000007&
      Height          =   1500
      Left            =   5700
      Picture         =   "frmMain.frx":360F
      ScaleHeight     =   100
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   31
      Top             =   5220
      Width           =   1500
      Begin VB.Image imgGuy 
         Height          =   360
         Left            =   630
         Top             =   570
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image img 
         Height          =   510
         Index           =   8
         Left            =   975
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
         Index           =   6
         Left            =   15
         Top             =   15
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
         Index           =   4
         Left            =   495
         Top             =   495
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
         Index           =   2
         Left            =   975
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
         Index           =   0
         Left            =   15
         Top             =   975
         Width           =   510
      End
   End
   Begin VB.PictureBox picHP 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   120
      Left            =   7305
      MouseIcon       =   "frmMain.frx":473C
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":5106
      ScaleHeight     =   8
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   90
      TabIndex        =   29
      ToolTipText     =   "Restore"
      Top             =   2070
      Width           =   1350
   End
   Begin VB.Frame fraMain 
      Height          =   4995
      Left            =   7260
      TabIndex        =   8
      Top             =   -75
      Width           =   1440
      Begin VB.PictureBox picCha 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
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
         ForeColor       =   &H80000008&
         Height          =   1350
         Left            =   45
         ScaleHeight     =   88
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   88
         TabIndex        =   5
         Top             =   735
         Width           =   1350
      End
      Begin VB.Image imgSkill 
         Height          =   150
         Left            =   990
         Picture         =   "frmMain.frx":541B
         Top             =   3990
         Width           =   225
      End
      Begin VB.Label lblCha 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Character"
         Height          =   225
         Left            =   60
         TabIndex        =   24
         Top             =   195
         Width           =   1350
      End
      Begin VB.Line Line6 
         BorderColor     =   &H80000014&
         X1              =   60
         X2              =   1360
         Y1              =   2565
         Y2              =   2565
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000010&
         X1              =   60
         X2              =   1360
         Y1              =   2550
         Y2              =   2550
      End
      Begin VB.Label lblAlias 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "- Alias -"
         Height          =   225
         Left            =   60
         TabIndex        =   30
         Top             =   465
         Width           =   1350
      End
      Begin VB.Label lblEP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Experience Points"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   90
         TabIndex        =   28
         Top             =   4305
         Width           =   1290
      End
      Begin VB.Label lblExp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   600
         TabIndex        =   27
         Top             =   4575
         Width           =   180
      End
      Begin VB.Label lblLevel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   1260
         TabIndex        =   26
         Top             =   2295
         Width           =   90
      End
      Begin VB.Label Level 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   60
         TabIndex        =   25
         Top             =   2295
         Width           =   390
      End
      Begin VB.Label lblSta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mana"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   6
         Left            =   60
         TabIndex        =   23
         Top             =   3960
         Width           =   420
      End
      Begin VB.Label lblSta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dexterity"
         Height          =   225
         Index           =   5
         Left            =   60
         TabIndex        =   22
         Top             =   3735
         Width           =   690
      End
      Begin VB.Label lblSta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agility"
         Height          =   225
         Index           =   4
         Left            =   60
         TabIndex        =   21
         Top             =   3510
         Width           =   510
      End
      Begin VB.Label lblSta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Toughness"
         Height          =   225
         Index           =   3
         Left            =   60
         TabIndex        =   20
         Top             =   3285
         Width           =   780
      End
      Begin VB.Label lblSta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Strengh"
         Height          =   225
         Index           =   2
         Left            =   60
         TabIndex        =   19
         Top             =   3060
         Width           =   540
      End
      Begin VB.Label lblSta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perception"
         Height          =   225
         Index           =   1
         Left            =   60
         TabIndex        =   18
         Top             =   2835
         Width           =   780
      End
      Begin VB.Label lblSta 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Knowledge"
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   17
         Top             =   2610
         Width           =   810
      End
      Begin VB.Label lblVal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   6
         Left            =   1260
         TabIndex        =   16
         Top             =   3960
         Width           =   90
      End
      Begin VB.Label lblVal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   5
         Left            =   1260
         TabIndex        =   15
         Top             =   3735
         Width           =   90
      End
      Begin VB.Label lblVal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   4
         Left            =   1260
         TabIndex        =   14
         Top             =   3510
         Width           =   90
      End
      Begin VB.Label lblVal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   3
         Left            =   1260
         TabIndex        =   13
         Top             =   3285
         Width           =   90
      End
      Begin VB.Label lblVal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   1260
         TabIndex        =   12
         Top             =   3060
         Width           =   90
      End
      Begin VB.Label lblVal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   1260
         TabIndex        =   11
         Top             =   2835
         Width           =   90
      End
      Begin VB.Label lblVal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   1260
         TabIndex        =   10
         Top             =   2610
         Width           =   90
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000010&
         X1              =   60
         X2              =   1360
         Y1              =   4245
         Y2              =   4245
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000014&
         X1              =   60
         X2              =   1360
         Y1              =   4260
         Y2              =   4260
      End
   End
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
      Height          =   615
      ItemData        =   "frmMain.frx":54A4
      Left            =   30
      List            =   "frmMain.frx":54A6
      TabIndex        =   9
      Top             =   5220
      Width           =   5625
   End
   Begin VB.Frame fraBar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   30
      TabIndex        =   6
      Top             =   5805
      Width           =   5640
      Begin VB.Image imgBack 
         Height          =   780
         Left            =   45
         Picture         =   "frmMain.frx":54A8
         Top             =   120
         Width           =   5580
      End
      Begin VB.Image imgBar 
         Height          =   720
         Index           =   6
         Left            =   4800
         Picture         =   "frmMain.frx":5DED
         Tag             =   "Examine "
         Top             =   120
         Width           =   720
      End
      Begin VB.Image imgBar 
         Height          =   720
         Index           =   5
         Left            =   4005
         Picture         =   "frmMain.frx":6051
         Tag             =   "Close "
         Top             =   120
         Width           =   720
      End
      Begin VB.Image imgBar 
         Height          =   720
         Index           =   1
         Left            =   825
         Picture         =   "frmMain.frx":61EF
         Tag             =   "Read "
         Top             =   120
         Width           =   720
      End
      Begin VB.Image imgBar 
         Height          =   720
         Index           =   4
         Left            =   3195
         Picture         =   "frmMain.frx":63AE
         Tag             =   "Put "
         Top             =   120
         Width           =   720
      End
      Begin VB.Image imgBar 
         Height          =   720
         Index           =   3
         Left            =   2415
         Picture         =   "frmMain.frx":65D5
         Tag             =   "Attack to "
         Top             =   120
         Width           =   720
      End
      Begin VB.Image imgBar 
         Height          =   720
         Index           =   2
         Left            =   1635
         Picture         =   "frmMain.frx":679B
         Tag             =   "Talk to "
         Top             =   120
         Width           =   720
      End
      Begin VB.Image imgBar 
         Height          =   720
         Index           =   0
         Left            =   45
         Picture         =   "frmMain.frx":6911
         Top             =   120
         Width           =   720
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000014&
         X1              =   3990
         X2              =   3990
         Y1              =   135
         Y2              =   835
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         X1              =   3975
         X2              =   3975
         Y1              =   135
         Y2              =   835
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   1620
         X2              =   1620
         Y1              =   140
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   1605
         X2              =   1605
         Y1              =   140
         Y2              =   840
      End
      Begin VB.Label lblCop 
         Alignment       =   2  'Center
         Caption         =   "Open "
         Height          =   225
         Left            =   4035
         TabIndex        =   32
         Top             =   600
         Width           =   720
      End
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
      ItemData        =   "frmMain.frx":6B4A
      Left            =   4830
      List            =   "frmMain.frx":6B51
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   4230
      Visible         =   0   'False
      Width           =   2250
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
      ItemData        =   "frmMain.frx":6B5D
      Left            =   2430
      List            =   "frmMain.frx":6B64
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   4230
      Visible         =   0   'False
      Width           =   2250
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
      IntegralHeight  =   0   'False
      ItemData        =   "frmMain.frx":6B70
      Left            =   30
      List            =   "frmMain.frx":6B77
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   4230
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.TextBox txtMain 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   4875
      Left            =   30
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmMain.frx":6B83
      Top             =   30
      Width           =   7200
   End
   Begin VB.Image imgOk 
      Height          =   330
      Left            =   5100
      Picture         =   "frmMain.frx":6BB4
      Top             =   4890
      Width           =   525
   End
   Begin VB.Image imgAni 
      Height          =   360
      Index           =   3
      Left            =   8430
      Picture         =   "frmMain.frx":6C5E
      Top             =   4950
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAni 
      Height          =   360
      Index           =   2
      Left            =   8040
      Picture         =   "frmMain.frx":6D8B
      Top             =   4950
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAni 
      Height          =   360
      Index           =   1
      Left            =   7650
      Picture         =   "frmMain.frx":6EBB
      Top             =   4950
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgAni 
      Height          =   360
      Index           =   0
      Left            =   7260
      Picture         =   "frmMain.frx":6FE9
      Top             =   4950
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgLook 
      Height          =   450
      Left            =   7770
      Picture         =   "frmMain.frx":711A
      Top             =   5730
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgUp 
      Height          =   450
      Left            =   7770
      Picture         =   "frmMain.frx":7512
      Top             =   5730
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgDown 
      Height          =   450
      Left            =   7770
      Picture         =   "frmMain.frx":790A
      Top             =   5730
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgArrow 
      Height          =   450
      Index           =   5
      Left            =   8220
      Picture         =   "frmMain.frx":7AA6
      Tag             =   ">Move East."
      Top             =   5730
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   5700
      TabIndex        =   7
      Top             =   4950
      Width           =   1500
   End
   Begin VB.Image imgArrow 
      Height          =   450
      Index           =   1
      Left            =   7770
      Picture         =   "frmMain.frx":7B5E
      Tag             =   ">Move South."
      Top             =   6180
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgArrow 
      Height          =   450
      Index           =   7
      Left            =   7770
      Picture         =   "frmMain.frx":7C1F
      Tag             =   ">Move North."
      Top             =   5280
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgArrow 
      Height          =   450
      Index           =   3
      Left            =   7320
      Picture         =   "frmMain.frx":7CDE
      Tag             =   ">Move West."
      Top             =   5730
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Image imgRing 
      Height          =   1350
      Left            =   7320
      Picture         =   "frmMain.frx":7D96
      Top             =   5280
      Width           =   1350
   End
   Begin VB.Menu mnuFil 
      Caption         =   "&File"
      Begin VB.Menu mnuFilNew 
         Caption         =   "&New Quest..."
      End
      Begin VB.Menu mnuFilBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilExt 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuVie 
      Caption         =   "&View"
      Begin VB.Menu mnuVieHal 
         Caption         =   "&Hall of Fame..."
      End
   End
   Begin VB.Menu mnuToo 
      Caption         =   "&Tools"
      Begin VB.Menu mnuTooMak 
         Caption         =   "Quest-Maker..."
      End
   End
   Begin VB.Menu mnuHel 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelAbo 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuInv 
      Caption         =   "Inventory"
      Visible         =   0   'False
      Begin VB.Menu mnuInvWea 
         Caption         =   "Wear"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuInvBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInvObj 
         Caption         =   "Empty"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu mnuInvBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInvPic 
         Caption         =   "Pick Up"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strR As String, mnu As Menu, r%, c%, Rst$, Task As Phrase
Private Type Phrase
  Action As String
  Subject As String
  Use As String
End Type

Private Sub Form_Load()
  Load frmMaker
  Me.Show
End Sub

Private Sub mnuFilNew_Click()
  frmNew.Show
End Sub

Private Sub mnuVieHal_Click()
  frmHall.Show
End Sub

Private Sub mnuTooMak_Click()
  frmMaker.Show
  Hide
End Sub

Private Sub mnuHelAbo_Click()
  frmAbout.Show
End Sub

Private Sub mnuInvObj_Click(Index As Integer)
  txtEntry = "Drop " & mnuInvObj(Index).Caption
End Sub

Private Sub mnuInvWea_Click()
  txtEntry = "Wear "
End Sub

Private Sub mnuInvPic_Click()
  txtEntry = "Pick Up "
End Sub

Private Sub mnuFilExt_Click()
  End
End Sub

Private Sub tmrAni_Timer()
  imgGuy.Picture = imgAni(0).Picture
  tmrAni.Enabled = False
End Sub

Private Sub txtMain_Change()
  txtMain.SelStart = Len(txtMain)
End Sub

Private Sub txtMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgBar_MouseMove 7, Button, Shift, X, Y
End Sub

Private Sub imgBar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  For i% = 0 To 6
    imgBar(i%).BorderStyle = IIf(Index = i%, 1, 0)
  Next i%
End Sub

Private Sub picHP_Click()
  Dim iDiff As Integer
  
  If picHP.Width < 90 And lblExp > 0 Then iDiff = 90 - picHP.Width Else Exit Sub
  If iDiff > lblExp Then picHP.Width = picHP.Width + lblExp: lblExp = 0 Else picHP.Width = 90: lblExp = lblExp - iDiff
  txtMain = txtMain & vbCrLf & vbCrLf & ">Health." & vbCrLf & "( " & Val(picHP.Width / 0.9) & "% Health )"
End Sub

Private Sub imgLook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lstSel.Clear: txtEntry = ""
  imgLook.Picture = imgDown.Picture
  txtMain = txtMain & Replace("||>Look|" & ReadINI("Description", strR, strFree), "|", vbCrLf)
  If lstTmp(0).ListCount + lstTmp(1).ListCount + lstTmp(2).ListCount = 0 And ReadINI("Description", strR, strFree) = "" Then txtMain = txtMain & " This place has been thoroughly searched"
  
  For i% = 0 To lstTmp(0).ListCount - 1
    If PStr("Find", ReadINI("Properties", lstTmp(0).List(i%), strFree), "Return", "index") = "0" And PStr("Find", ReadINI("Properties", lstTmp(0).List(i%), strFree), , "after") = 0 Then
      If i% = 0 Then varTmp = " Also here is a " Else varTmp = " and a "
      txtMain = txtMain & varTmp & lstTmp(0).List(i%)
    End If
  Next i%
  
  For i% = 0 To lstTmp(1).ListCount - 1
    If PStr("Find", ReadINI("Properties", lstTmp(1).List(i%), strFree), , "after") = 0 Then txtMain = txtMain & " " & lstTmp(1).List(i%) & " is here. "
  Next i%
End Sub

Private Sub imgLook_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgLook.Picture = imgUp.Picture
End Sub

Private Sub LoadQuest()
  For i% = 0 To 2
    lstTmp(i%).Clear
  Next i%
  
  lblTitle = PStr("Find", ReadINI("Map", "Title", strFree), "Return", strR)
  txtMain = Replace(txtMain & "|" & ReadINI("Description", strR, strFree) & ReadINI("Intro", strR, strFree), "|", vbCrLf)
  varTmp = ReadINI("Movement", strR, strFree)
  If varTmp <> "" Then SplitAndAdd lstTmp(1), varTmp
  varTmp = Split(ReadINI("Element", strR, strFree), "|")
  If UBound(varTmp) > -1 Then SplitAndAdd lstTmp(0), varTmp(0): SplitAndAdd lstTmp(1), varTmp(1): SplitAndAdd lstTmp(2), varTmp(2)
  varTmp = Split(ReadINI("Element", "Actions", strFree), "|")
  For i% = 1 To UBound(varTmp)
    lstAct.AddItem Split(varTmp(i%), "")(0)
  Next i%
  
  For Each mnu In mnuInvObj
    If mnu.Index > 0 Then GetIndex lstTmp(0), mnu.Caption, "Remove": lstTmp(0).AddItem mnu.Caption
  Next mnu
End Sub

Private Sub SaveQuest()
  Me.Tag = ""
  For i% = 0 To 2
    For e% = 0 To lstTmp(i%).ListCount - 1
      lstTmp(i%).Tag = lstTmp(i%).Tag & "" & lstTmp(i%).List(e%)
    Next e%
    Me.Tag = Me.Tag & lstTmp(i%).Tag & "|"
    lstTmp(i%).Tag = ""
  Next i%
  
  WriteINI "Intro", strR, "", strFree
  WriteINI "Movement", strR, "", strFree
  WriteINI "Element", strR, Me.Tag, strFree
End Sub

Private Sub ShowMap(bCase As Boolean)
  Dim iNum As Integer, iTile As Integer
  
  varTmp = ReadINI("Map", "Title", strFree)
  For i% = c% - 1 To c% + 1
    For e% = r% - 1 To r% + 1
      If InStr(varTmp, "|" & e% & "," & i% & "") = 0 Then img(iNum).Visible = False Else img(iNum).Visible = bCase
      If i% = c% Xor e% = r% Then If img(iNum).Visible = False Then imgArrow(iNum).Visible = False Else imgArrow(iNum).Visible = bCase
      If InStr(varTmp, "|" & e% - 1 & "," & i% & "") > 0 Then iTile = 1 Else iTile = 0
      If InStr(varTmp, "|" & e% & "," & i% + 1 & "") > 0 Then iTile = iTile + 2
      If InStr(varTmp, "|" & e% + 1 & "," & i% & "") > 0 Then iTile = iTile + 4
      If InStr(varTmp, "|" & e% & "," & i% - 1 & "") > 0 Then iTile = iTile + 8
      img(iNum).Picture = imlTiles.ListImages(iTile + 1).Picture
      iNum = iNum + 1
    Next e%
  Next i%
End Sub

Private Sub imgArrow_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim strTmp As String
  
  lstSel.Clear: txtEntry = ""
  tmrAni.Enabled = False
  imgArrow(Index).Visible = False
  Rst$ = strR & "," & CStr(Index \ 2)
  txtMain = txtMain & vbCrLf & vbCrLf & imgArrow(Index).Tag
  strTmp = PStr("Find", ReadINI("Restriction", "Room", strFree), "Return", Rst$)
  If strTmp <> "" Then txtMain = txtMain & vbCrLf & ReadINI("Restriction", Rst$, strFree): Exit Sub
  SaveQuest
  If Index = 7 Then c% = c% + 1
  If Index = 5 Then r% = r% + 1
  If Index = 3 Then r% = r% - 1
  If Index = 1 Then c% = c% - 1
  strR = r% & "," & c%
  imgGuy.Picture = imgAni(Index \ 2).Picture
  LoadQuest
End Sub

Private Sub imgArrow_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  tmrAni.Enabled = True
  ShowMap True
End Sub

Private Sub imgBar_Click(Index As Integer)
  mnuInvPic.Enabled = IIf(mnuInvObj.Count = 6, False, True)
  mnuInvWea.Enabled = IIf(mnuInvObj.Count > 1, True, False)
  mnuInvObj(0).Visible = IIf(mnuInvObj.Count > 1, False, True)
  If Index = 0 Then txtEntry = "": lstSel.Clear: PopupMenu mnuInv, vbPopupMenuLeftAlign, fraBar.Left, fraBar.Top + fraBar.Height Else txtEntry = imgBar(Index).Tag
  If Index = 5 Then txtEntry = lblCop: lblCop = IIf(lblCop = "Open ", "Close ", "Open ")
  ExecuteTask Trim(txtEntry)
End Sub

Private Function Parse(strPhr As String) As Phrase
  Parse.Action = "": Parse.Subject = "": Parse.Use = ""
  For i% = 0 To lstAct.ListCount - 1
    If InStr(1, strPhr, lstAct.List(i%), vbTextCompare) > 0 Then Parse.Action = lstAct.List(i%): i% = lstAct.ListCount - 1
  Next i%
  
  For a% = 0 To lstTmp(1).ListCount - 1
    If InStr(1, strPhr, lstTmp(1).List(a%), vbTextCompare) > 0 Then
      Parse.Subject = lstTmp(1).List(a%): strPhr = Replace(strPhr, Parse.Subject, "")
    End If
  Next a%
  
  For e% = 0 To 2 Step 2
    For a% = 0 To lstTmp(e%).ListCount - 1
      If InStr(1, strPhr, lstTmp(e%).List(a%), vbTextCompare) > 0 Then
        If Parse.Subject = "" Then Parse.Subject = lstTmp(e%).List(a%): strPhr = Replace(strPhr, Parse.Subject, "") Else Parse.Use = lstTmp(e%).List(a%)
      End If
    Next a%
  Next e%
End Function

Private Sub lstSel_Click()
  ExecuteTask Trim(txtEntry) & " " & lstSel.Text
End Sub

Private Sub txtEntry_KeyPress(KeyAscii As Integer)
  If Int(KeyAscii) = 13 Then ExecuteTask Trim(txtEntry)
End Sub

Private Sub ExecuteTask(strPhr As String)
  Dim strAct As String, strPro As String, strUse As String, strTsk As String, strTmp As String, strMsg As String
  Dim strText As String
  Static iMenu As Integer
  
  Randomize
  lstSel.Clear
  strText = strPhr
  Task = Parse(strText)
  
  strUse = ReadINI("Properties", Task.Use, strFree)
  strAct = ReadINI("Actions", Task.Subject, strFree)
  strPro = ReadINI("Properties", Task.Subject, strFree)
  strTsk = ReadINI("Properties", Task.Subject & "/" & Task.Action, strFree)

  If Task.Action <> "ask" And Task.Action <> "talk" Then If PStr("Find", strUse, , "after") > 0 Then strMsg = "You don't have a " & Task.Use & " . ": GoTo Message
  If PStr("Find", strPro, , "after") > 0 Then strMsg = "You don't see that here": GoTo Message
  If PStr("Find", strTsk, , "use") > 0 Then If LCase(PStr("Find", strTsk, "Return", "use")) <> LCase(Task.Use) Then strMsg = "An element it's required to perform this task": GoTo Message
  If PStr("Find", strTsk, , "dice") > 0 Then frmRoll.Show: GoTo Message
  
  Select Case Task.Action
    Case "pick", "take", "pick up"
      For i% = 0 To lstTmp(0).ListCount - 1
        strTmp = ReadINI("Properties", lstTmp(0).List(i%), strFree)
        If PStr("Find", ReadINI("Actions", lstTmp(0).List(i%), strFree), , Task.Action) > 0 Or PStr("Find", strTmp, , "stationary") = 0 Then If PStr("Find", strTmp, "Return", "container") <> "player" And PStr("Find", strTmp, , "after") = 0 Then lstSel.AddItem lstTmp(0).List(i%)
      Next i%
      
      If Task.Subject = "" Then Exit Sub
      If PStr("Find", strAct, , Task.Action) = 0 And PStr("Find", strPro, , "stationary") > 0 Then strMsg = "Nothing happened": GoTo Message
      If PStr("Find", strPro, "Return", "container") = "player" Then strMsg = "You are carring the " & Task.Subject: GoTo Message

      iMenu = iMenu + 1
      Load mnuInvObj(iMenu)
      mnuInvObj(iMenu).Visible = True
      mnuInvObj(iMenu).Enabled = True
      mnuInvObj(iMenu).Caption = Task.Subject
      strPro = PStr("Add", strPro, 0, "index")
      strPro = PStr("Add", strPro, "player", "container")
      strPro = PStr("Add", strPro, Task.Subject, "parent")
      WriteINI "Properties", Task.Subject, strPro, strFree
      If PStr("Find", strAct, "Return", Task.Action) = "" Then strMsg = "You " & Task.Action & " up a " & Task.Subject
      
    Case "drop", "put down"
      If mnuInvObj(0).Visible = True Then strMsg = "You are carring nothing": GoTo Message
      For Each mnu In mnuInvObj
        If mnu.Index > 0 Then lstSel.AddItem mnu.Caption
      Next mnu

      If Task.Subject = "" Then Exit Sub
      If PStr("Find", strPro, "Return", "container") <> "player" Then strMsg = "You are not carring a " & Task.Subject: GoTo Message
      If PStr("Find", strAct, "Return", Task.Action) = "" Then strMsg = "You " & Task.Action & " the " & Task.Subject
      
      For Each mnu In mnuInvObj
        If Task.Subject = mnu.Caption Then
          If mnu.Checked = True Then lblVal(PStr("Find", strPro, "Return", "attribute")) = CInt(lblVal(PStr("Find", strPro, "Return", "attribute"))) - PStr("Find", strPro, "Return", "size") - 1
          Unload mnu
        End If
      Next mnu
      WriteINI "Properties", Task.Subject, PStr("Remove", strPro, "", "container"), strFree
      
    Case "wear", "put on"
      If mnuInvObj(0).Visible = True Then strMsg = "You are carring nothing": GoTo Message
      For Each mnu In mnuInvObj
        strTmp = ReadINI("Properties", mnu.Caption, strFree)
        If PStr("Find", strTmp, , "wearable") > 0 And mnu.Index > 0 And mnu.Checked = False Then lstSel.AddItem mnu.Caption
      Next mnu

      If Task.Subject = "" Then Exit Sub
      If PStr("Find", strAct, , Task.Action) = 0 And PStr("Find", strPro, , "wearable") = 0 Then strMsg = "Nothing happened": GoTo Message
      If PStr("Find", strPro, "Return", "container") <> "player" Then strMsg = "You are not carring a " & Task.Subject: GoTo Message
      If PStr("Find", strPro, , "attribute") = 0 Then WriteINI "Properties", Task.Subject, PStr("Add", strPro, CInt(Rnd * 5), "attribute"), strFree
      If PStr("Find", strAct, "Return", Task.Action) = "" Then strMsg = "You " & Task.Action & " the " & Task.Subject
      For Each mnu In mnuInvObj
        If Task.Subject = mnu.Caption Then If mnu.Checked = False Then mnu.Checked = True Else strMsg = "You are wearing the " & Task.Subject: GoTo Message
      Next mnu
   
    Case "talk", "ask"
      For i% = 0 To lstTmp(1).ListCount - 1
        strTmp = ReadINI("Properties", lstTmp(1).List(i%), strFree)
        If PStr("Find", strTmp, , "after") = 0 Then lstSel.AddItem lstTmp(1).List(i%)
      Next i%

      If Task.Subject = "" Then Exit Sub
      If Task.Use <> "" Then
        strTmp = ReadINI("Actions", Task.Subject & "/ask", strFree)
        If PStr("Find", strTmp, "Return", Task.Use) = "" Then strMsg = "You can't talk to that": GoTo Message
        strMsg = Task.Subject & " - " & PStr("Find", strTmp, "Return", Task.Use)
        strTmp = "DoAfter": GoTo Message
      End If
      
      lstSel.Clear
      varTmp = Split(ReadINI("Actions", Task.Subject & "/ask", strFree), "|")
      For i% = 1 To UBound(varTmp)
        lstSel.AddItem Split(varTmp(i%), "")(0)
      Next i%
      txtEntry = Task.Action & " to " & Task.Subject & " about": Exit Sub

    Case "open", "close"
      For i% = 0 To lstTmp(0).ListCount - 1
        strTmp = ReadINI("Properties", lstTmp(0).List(i%), strFree)
        If PStr("Find", ReadINI("Actions", lstTmp(0).List(i%), strFree), , Task.Action) > 0 Or PStr("Find", strTmp, "Return", "openable") = IIf(Task.Action = "open", "closed", "opened") Then If PStr("Find", strTmp, , "after") = 0 Then lstSel.AddItem lstTmp(0).List(i%)
      Next i%
      
      If Task.Subject = "" Then Exit Sub
      If PStr("Find", strAct, , Task.Action) = 0 And PStr("Find", strPro, , "openable") = 0 Then strMsg = "Nothing happened": GoTo Message
      If PStr("Find", strPro, , "openable") > 0 Then WriteINI "Properties", Task.Subject, PStr("Add", strPro, IIf(Task.Action = "open", "opened", "closed"), "openable"), strFree
      If PStr("Find", strAct, "Return", Task.Action) = "" Then If PStr("Find", strPro, "Return", "openable") = IIf(Task.Action = "open", "opened", "closed") Then strMsg = "The " & Task.Subject & " is already " & IIf(Task.Action = "open", "opened", "closed") Else strMsg = "The " & Task.Subject & " is " & IIf(Task.Action = "open", "opened", "closed")

    Case "put"
      For i% = 0 To lstTmp(0).ListCount - 1
        strTmp = ReadINI("Properties", lstTmp(0).List(i%), strFree)
        If PStr("Find", strTmp, , "stationary") = 0 Then If PStr("Find", strTmp, , "after") = 0 Then lstSel.AddItem lstTmp(0).List(i%)
      Next i%
      
      If Task.Subject = "" Then Exit Sub
      If PStr("Find", strAct, , Task.Action) = 0 And PStr("Find", strPro, , "stationary") > 0 Then strMsg = "Nothing happened": GoTo Message
      
      If Task.Use <> "" Then
        If PStr("Find", strPro, "Return", "size") > PStr("Find", strUse, "Return", "size") Then strMsg = "The " & Task.Subject & " is bigger than the " & Task.Use: GoTo Message
        If PStr("Find", strUse, "Return", "openable") = "closed" Then strMsg = "The " & Task.Use & " is closed": GoTo Message
        If PStr("Find", strUse, , "openable") = 0 Then strMsg = "The " & Task.Use & " is not a container": GoTo Message
        If PStr("Find", strAct, "Return", Task.Action) = "" Then strMsg = "The " & Task.Subject & " is into the " & Task.Use
        strTmp = "DoAfter": GoTo Message
      End If
      
      lstSel.Clear
      For i% = 0 To lstTmp(0).ListCount - 1
        strTmp = ReadINI("Properties", lstTmp(0).List(i%), strFree)
        If lstTmp(0).List(i%) <> Task.Subject And PStr("Find", strTmp, , "openable") > 0 And PStr("Find", strTmp, , "after") = 0 Then lstSel.AddItem lstTmp(0).List(i%)
      Next i%
      txtEntry = "Put " & Task.Subject & " into ": Exit Sub
      
    Case "examine", "look"
      For e% = 0 To 2
        For i% = 0 To lstTmp(e%).ListCount - 1
          strTmp = ReadINI("Properties", lstTmp(e%).List(i%), strFree)
          If PStr("Find", strTmp, , "after") = 0 Then lstSel.AddItem lstTmp(e%).List(i%)
        Next i%
      Next e%
      
      If Task.Subject = "" Then Exit Sub
      If PStr("Find", strAct, , "examine") > 0 Then Task.Action = "examine" Else Task.Action = "look" '1
      If PStr("Find", strAct, "Return", Task.Action) = "" Then strMsg = "Nothing special" '2
    
    Case Else
      For a% = 0 To 2
        For i% = 0 To lstTmp(a%).ListCount - 1
          strTmp = ReadINI("Actions", lstTmp(a%).List(i%), strFree)
          If PStr("Find", strTmp, , Task.Action) > 0 And PStr("Find", ReadINI("Properties", lstTmp(a%).List(i%), strFree), , "after") = 0 Then lstSel.AddItem lstTmp(a%).List(i%)
        Next i%
      Next a%
      
      If Task.Subject = "" Then Exit Sub
      If PStr("Find", strAct, , Task.Action) = 0 Then strMsg = "Nothing happened": GoTo Message
  End Select
  If ReadINI("Ending", "Final", strFree) = Task.Subject & "/" & Task.Action Then frmFinal.Show: FormInit False: Exit Sub
  strTmp = "DoAfter"
  
Message:
  lstSel.Clear: txtEntry = ""
  txtMain = txtMain & vbCrLf & vbCrLf & ">" & strPhr & IIf(strMsg <> "", vbCrLf & strMsg, "")
  If strTmp = "DoAfter" Then AfterTask
End Sub

Private Sub imgOk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgOk.BorderStyle = 1
  ExecuteTask txtEntry
End Sub

Private Sub imgOk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgOk.BorderStyle = 0
End Sub

Public Sub Update(iSum As Integer)
  Dim strMsg As String
  Select Case iSum
    Case Is > 0
      AfterTask
      lblExp = lblExp + iSum
      strMsg = "( +" & iSum & " Exp. Points )"
      If lblLevel = (lblExp \ 50) Then lblLevel = lblLevel + 1: lblVal(6) = lblVal(6) + 1

    Case Is < 0
      strMsg = "Failed Intent." & vbCrLf & "( " & iSum & " Hit Points )"
      If picHP.Width + iSum < 1 Then picHP.Width = 0: FormInit False: txtMain = vbCrLf & "You start to feel weak...death has found you!!!" & vbCrLf & "The End": Exit Sub
      picHP.Width = picHP.Width + iSum
  End Select
  txtMain = txtMain & vbCrLf & strMsg
End Sub

Private Sub AfterTask()
  Dim strAct As String, strPro As String, strUse As String, strTac As String, strTsk As String, strTmp As String
  
  strUse = ReadINI("Properties", Task.Use, strFree)
  strAct = ReadINI("Actions", Task.Subject, strFree)
  strPro = ReadINI("Properties", Task.Subject, strFree)
  strTac = ReadINI("Actions", Task.Subject & "/" & Task.Action, strFree)
  strTsk = ReadINI("Properties", Task.Subject & "/" & Task.Action, strFree)
  
  strTmp = PStr("Find", strAct, "Return", Task.Action)
  txtMain = txtMain & IIf(strTmp <> "", vbCrLf & strTmp, "")
  
  strTsk = PStr("Remove", strTsk, "", "dice")
  If PStr("Find", strTac, , "permanent") > 0 Then Else strAct = PStr("Remove", strAct, , Task.Action)
  If PStr("Find", strTsk, , "hide") > 0 Then strPro = PStr("Add", strPro, "hide", "after")
  For Each mnu In mnuInvObj
    If PStr("Find", strTsk, , "hide") > 0 And mnu.Caption = Task.Subject Then Unload mnu
  Next mnu
  If PStr("Find", strTsk, "Return", "Ndes") <> "" Then strAct = PStr("Add", strAct, PStr("Find", strTsk, "Return", "Ndes"), "look")
  If PStr("Find", strTsk, "Return", "stationary") = "yes" Then strPro = PStr("Add", strPro, , "stationary")
  If PStr("Find", strTsk, "Return", "stationary") = "no" Then strPro = PStr("Remove", strPro, , "stationary")

  For i% = 0 To lstTmp(0).ListCount - 1
    strTmp = ReadINI("Properties", lstTmp(0).List(i%), strFree)
    If PStr("Find", strTmp, "Return", "container") = Task.Subject Then If PStr("Find", strTsk, "Return", "drop") = lstTmp(0).List(i%) Or PStr("Find", strTsk, "Return", "drop") = "<All>" Then strTmp = PStr("Remove", strTmp, , "after"): strTmp = PStr("Remove", strTmp, , "container"): strTmp = PStr("Remove", strTmp, , "parent"): strTmp = PStr("Add", strTmp, 0, "index")
    WriteINI "Properties", lstTmp(0).List(i%), strTmp, strFree
  Next i%
  
  strTmp = ReadINI("Restriction", "Room", strFree)
  For i% = 0 To 3
    If Task.Subject & "/" & Task.Action = PStr("Find", strTmp, "Return", strR & "," & i%) Then strTmp = PStr("Add", strTmp, "", strR & "," & i%)
  Next i%
  WriteINI "Restriction", "Room", strTmp, strFree
  
  For e% = 0 To 2
    For i% = 0 To lstTmp(e%).ListCount - 1
      strTmp = ReadINI("Properties", lstTmp(e%).List(i%), strFree)
      If PStr("Find", strTmp, "Return", "after") = Task.Subject & "/" & Task.Action Then strTmp = PStr("Remove", strTmp, "", "after")
      If PStr("Find", strTmp, "Return", "index") > PStr("Find", strPro, "Return", "index") And PStr("Find", strTmp, "Return", "parent") = PStr("Find", strPro, "Return", "parent") Then
        If Task.Action = "close" Then If PStr("Find", strTmp, , "after") = 0 Then strTmp = PStr("Add", strTmp, Task.Subject & "/open", "after")
        If Task.Action = "put" Then strTmp = PStr("Add", strTmp, PStr("Find", strTmp, "Return", "index") + 1, "index")
        If Task.Action = "put" Then strTmp = PStr("Add", strTmp, Task.Use, "parent")
      End If
      WriteINI "Properties", lstTmp(e%).List(i%), strTmp, strFree
    Next i%
  Next e%

  If Task.Action = "put" Then
    strPro = PStr("Add", strPro, PStr("Find", strUse, "Return", "index") + 1, "index")
    strPro = PStr("Add", strPro, PStr("Find", strUse, "Return", "parent"), "parent")
    strPro = PStr("Add", strPro, Task.Use, "container")
  End If
  If Task.Action = "wear" Then
    lblVal(PStr("Find", strPro, "Return", "attribute")) = CInt(lblVal(PStr("Find", strPro, "Return", "attribute"))) + PStr("Find", strPro, "Return", "size") + 1
    txtMain = txtMain & vbCrLf & "( " & lblSta(PStr("Find", strPro, "Return", "attribute")) & " +" & (PStr("Find", strPro, "Return", "size") + 1) & " )"
  End If
  
  WriteINI "Actions", Task.Subject, strAct, strFree
  WriteINI "Properties", Task.Subject, strPro, strFree
  WriteINI "Properties", Task.Subject & "/" & Task.Action, strTsk, strFree
End Sub

Private Sub FormInit(bCase As Boolean)
  lstSel.Clear: txtEntry = ""
  imgBack.Top = IIf(bCase = True, 900, 120)
  txtEntry.Enabled = bCase
  imgLook.Visible = bCase
  imgGuy.Visible = bCase
  ShowMap bCase
  
  For Each mnu In mnuInvObj
    If mnu.Index > 0 Then Unload mnu
  Next mnu
End Sub

Public Sub Starting(strPath As String, strName As String)
  picHP.Width = 90
  lblLevel = 1
  lblExp = 0

On Error Resume Next
  strName = App.Path & "\Data\" & strName
  strFree = App.Path & "\Free.tmp"
  FileCopy strPath, strFree
  
  varTmp = Split(ReadINI("Starting", "Start", strFree), ",")
  r% = varTmp(0): c% = varTmp(1): strR = r% & "," & c%
  
  varTmp = Split(ReadINI("Character", "Data", strName), "")
  lblCha = varTmp(0): lblAlias = varTmp(1)
  lblSta(6) = varTmp(2): picCha.Tag = varTmp(3)
  picCha.Picture = LoadPicture(picCha.Tag)

  varTmp = Split(ReadINI("Character", "Stats", strName), "|")
  For i% = 0 To 6
    lblVal(i%) = varTmp(i%)
  Next i%
  
  txtMain = Replace(ReadINI("Starting", "Info", strFree), "|", vbCrLf)
  LoadQuest
  FormInit True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  mnuFilExt_Click
End Sub
