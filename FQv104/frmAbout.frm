VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "frmAbout"
   ClientHeight    =   4200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   280
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picAbout 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2100
      Left            =   2520
      ScaleHeight     =   140
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   1950
      Width           =   3000
      Begin VB.Timer tmrScroll 
         Interval        =   100
         Left            =   0
         Top             =   0
      End
      Begin VB.TextBox txtAbout 
         Alignment       =   2  'Center
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
         Height          =   1800
         Left            =   150
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   2
         Text            =   "frmAbout.frx":4594
         Top             =   1425
         Width           =   2700
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   2880
      Picture         =   "frmAbout.frx":465A
      Top             =   240
      Width           =   2085
   End
   Begin VB.Label lblUrl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "www.geocities.com/freequestrpg/"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Left            =   2790
      MouseIcon       =   "frmAbout.frx":4B16
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1650
      Width           =   2475
   End
   Begin VB.Image imgGem 
      Height          =   345
      Left            =   5160
      Picture         =   "frmAbout.frx":4E20
      Top             =   270
      Width           =   345
   End
   Begin VB.Image imgClose 
      Height          =   345
      Left            =   5160
      ToolTipText     =   "Close"
      Top             =   270
      Width           =   345
   End
   Begin VB.Label lblAbout 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Abel Antonio Ricaurte © 2001-2003  Version 1.0.4"
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
      Height          =   450
      Left            =   2520
      TabIndex        =   1
      Top             =   810
      Width           =   3000
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgGem.Visible = True
End Sub

Private Sub imgGem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgGem.Visible = False
End Sub

Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgGem.Visible = True
  Me.Visible = False
End Sub

Private Sub lblUrl_Click()
  Shell "explorer.exe http://www.geocities.com/freequestrpg/"
End Sub

Private Sub tmrScroll_Timer()
  If txtAbout.Top = -txtAbout.Height Then txtAbout.Top = picAbout.Height
  txtAbout.Top = txtAbout.Top - 1
End Sub
