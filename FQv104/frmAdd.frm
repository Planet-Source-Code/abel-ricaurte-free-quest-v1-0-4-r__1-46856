VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAdd 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6690
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   500
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   446
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Default         =   -1  'True
      Height          =   270
      Left            =   3150
      TabIndex        =   7
      Top             =   7200
      Width           =   1425
   End
   Begin VB.CommandButton cmdCnl 
      Caption         =   "Cancel"
      Height          =   270
      Left            =   4575
      TabIndex        =   6
      Top             =   7200
      Width           =   1425
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   7140
      Left            =   150
      TabIndex        =   8
      Top             =   15
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   12594
      _Version        =   393216
      Style           =   1
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   529
      TabCaption(0)   =   "Object"
      TabPicture(0)   =   "frmAdd.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label55"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtName(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmbAfter(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Character"
      TabPicture(1)   =   "frmAdd.frx":03ED
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label8"
      Tab(1).Control(1)=   "Label15"
      Tab(1).Control(2)=   "txtName(1)"
      Tab(1).Control(3)=   "Frame8"
      Tab(1).Control(4)=   "cmbAfter(1)"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Other"
      TabPicture(2)   =   "frmAdd.frx":0498
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label62"
      Tab(2).Control(1)=   "Label28"
      Tab(2).Control(2)=   "txtName(2)"
      Tab(2).Control(3)=   "Frame2"
      Tab(2).Control(4)=   "cmbAfter(2)"
      Tab(2).Control(5)=   "chkInv"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Task"
      TabPicture(3)   =   "frmAdd.frx":057C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtAct"
      Tab(3).Control(1)=   "txtSub"
      Tab(3).Control(2)=   "Frame9"
      Tab(3).Control(3)=   "Label9"
      Tab(3).Control(4)=   "Label12"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "Starting"
      TabPicture(4)   =   "frmAdd.frx":0679
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label23"
      Tab(4).Control(1)=   "Label24"
      Tab(4).Control(2)=   "Label40"
      Tab(4).Control(3)=   "Frame17"
      Tab(4).Control(4)=   "txtTitle"
      Tab(4).Control(5)=   "txtAuthor"
      Tab(4).Control(6)=   "cmbStart"
      Tab(4).ControlCount=   7
      TabCaption(5)   =   "Battle"
      TabPicture(5)   =   "frmAdd.frx":0726
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "chkBat"
      Tab(5).Control(1)=   "Frame13"
      Tab(5).ControlCount=   2
      TabCaption(6)   =   "Ending"
      TabPicture(6)   =   "frmAdd.frx":07BF
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Label10"
      Tab(6).Control(1)=   "Frame18"
      Tab(6).Control(2)=   "cmbAfter(6)"
      Tab(6).ControlCount=   3
      Begin VB.CheckBox chkInv 
         Alignment       =   1  'Right Justify
         Caption         =   "Hidden"
         Height          =   210
         Left            =   -71850
         TabIndex        =   92
         Top             =   540
         Width           =   1080
      End
      Begin VB.ComboBox cmbAfter 
         Height          =   315
         Index           =   6
         ItemData        =   "frmAdd.frx":0897
         Left            =   -70950
         List            =   "frmAdd.frx":0899
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   480
         Width           =   1800
      End
      Begin VB.ComboBox cmbAfter 
         Height          =   315
         Index           =   2
         ItemData        =   "frmAdd.frx":089B
         Left            =   -70920
         List            =   "frmAdd.frx":089D
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   69
         Top             =   6600
         Width           =   1800
      End
      Begin VB.ComboBox cmbAfter 
         Height          =   315
         Index           =   1
         ItemData        =   "frmAdd.frx":089F
         Left            =   -70920
         List            =   "frmAdd.frx":08A1
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   6600
         Width           =   1800
      End
      Begin VB.ComboBox cmbAfter 
         Height          =   315
         Index           =   0
         ItemData        =   "frmAdd.frx":08A3
         Left            =   4080
         List            =   "frmAdd.frx":08A5
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   6600
         Width           =   1800
      End
      Begin VB.TextBox txtAct 
         Height          =   315
         Left            =   -70950
         TabIndex        =   3
         Top             =   480
         Width           =   1800
      End
      Begin VB.Frame Frame2 
         Height          =   5550
         Left            =   -74820
         TabIndex        =   53
         Top             =   840
         Width           =   6000
         Begin VB.CheckBox chkDin 
            Alignment       =   1  'Right Justify
            Caption         =   "Dinamic"
            Height          =   210
            Left            =   150
            TabIndex        =   90
            Top             =   1980
            Width           =   1080
         End
         Begin VB.TextBox txtDesO 
            Height          =   1080
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   54
            Top             =   480
            Width           =   4650
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Caption         =   "P   r   o   p   e   r   t   i   e   s"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000015&
            Height          =   255
            Left            =   1050
            TabIndex        =   91
            Top             =   1680
            Width           =   4650
         End
         Begin VB.Label Label61 
            Alignment       =   2  'Center
            Caption         =   "D   e   s   c   r   i   p   t   i   o   n"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000015&
            Height          =   255
            Left            =   1050
            TabIndex        =   57
            Top             =   240
            Width           =   4650
         End
         Begin VB.Label Label42 
            Caption         =   "Show"
            Height          =   210
            Left            =   180
            TabIndex        =   56
            Top             =   540
            Width           =   750
         End
         Begin VB.Label Label35 
            Height          =   210
            Left            =   210
            TabIndex        =   55
            Top             =   3820
            Width           =   750
         End
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Index           =   2
         Left            =   -73800
         TabIndex        =   2
         Top             =   480
         Width           =   1800
      End
      Begin VB.TextBox txtSub 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   1800
      End
      Begin VB.Frame Frame9 
         Height          =   5550
         Left            =   -74820
         TabIndex        =   48
         Top             =   840
         Width           =   6000
         Begin VB.CheckBox chkDice 
            Alignment       =   1  'Right Justify
            Caption         =   "Roll Dice"
            Height          =   210
            Left            =   180
            TabIndex        =   94
            Top             =   5160
            Width           =   1080
         End
         Begin VB.CheckBox chkPer 
            Alignment       =   1  'Right Justify
            Caption         =   "Continuous "
            Height          =   210
            Left            =   3030
            TabIndex        =   93
            Top             =   5160
            Width           =   1125
         End
         Begin VB.OptionButton optDin 
            Alignment       =   1  'Right Justify
            Caption         =   "Dinamic"
            Height          =   210
            Left            =   4500
            TabIndex        =   89
            Top             =   2400
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton optSta 
            Alignment       =   1  'Right Justify
            Caption         =   "Stationary"
            Height          =   210
            Left            =   3000
            TabIndex        =   88
            Top             =   2400
            Width           =   1080
         End
         Begin VB.TextBox txtNdes 
            Height          =   675
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   86
            Top             =   2670
            Width           =   4650
         End
         Begin VB.CheckBox chkHide 
            Alignment       =   1  'Right Justify
            Caption         =   "Hide"
            Height          =   210
            Left            =   180
            TabIndex        =   84
            Top             =   2400
            Width           =   1080
         End
         Begin VB.ComboBox cmbDrop 
            Height          =   315
            ItemData        =   "frmAdd.frx":08A7
            Left            =   1050
            List            =   "frmAdd.frx":08AE
            Style           =   2  'Dropdown List
            TabIndex        =   81
            Top             =   3420
            Width           =   1800
         End
         Begin VB.ComboBox cmbUse 
            Height          =   315
            ItemData        =   "frmAdd.frx":08BA
            Left            =   3900
            List            =   "frmAdd.frx":08C1
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   240
            Width           =   1800
         End
         Begin VB.TextBox txtMsg 
            Height          =   1080
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   49
            Top             =   900
            Width           =   4650
         End
         Begin VB.Label Label13 
            Caption         =   "Description"
            Height          =   210
            Left            =   180
            TabIndex        =   87
            Top             =   2730
            Width           =   810
         End
         Begin VB.Label Label3 
            Caption         =   "Drop"
            Height          =   210
            Left            =   180
            TabIndex        =   85
            Top             =   3480
            Width           =   810
         End
         Begin VB.Label Label36 
            Caption         =   "Use/Carry"
            Height          =   210
            Left            =   3000
            TabIndex        =   63
            Top             =   300
            Width           =   810
         End
         Begin VB.Label Label38 
            Alignment       =   2  'Center
            Caption         =   "S   u   b   j   e   c   t"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000015&
            Height          =   255
            Left            =   1050
            TabIndex        =   60
            Top             =   2100
            Width           =   4650
         End
         Begin VB.Label Label57 
            Alignment       =   2  'Center
            Caption         =   "M   e   s   s   a   g   e"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000015&
            Height          =   255
            Left            =   1050
            TabIndex        =   51
            Top             =   660
            Width           =   4650
         End
         Begin VB.Label Label16 
            Caption         =   "Show"
            Height          =   210
            Left            =   180
            TabIndex        =   50
            Top             =   960
            Width           =   810
         End
      End
      Begin VB.ComboBox cmbStart 
         Height          =   315
         Left            =   -70920
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   6600
         Width           =   1800
      End
      Begin VB.TextBox txtAuthor 
         Height          =   315
         Left            =   -73800
         TabIndex        =   43
         Text            =   "Anonymous"
         Top             =   900
         Width           =   4650
      End
      Begin VB.TextBox txtTitle 
         Height          =   315
         Left            =   -73800
         TabIndex        =   4
         Text            =   "Untitled"
         Top             =   480
         Width           =   4650
      End
      Begin VB.Frame Frame17 
         Height          =   5070
         Left            =   -74820
         TabIndex        =   40
         Top             =   1320
         Width           =   6000
         Begin VB.TextBox txtIntro 
            Height          =   4380
            Left            =   300
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   41
            Top             =   480
            Width           =   5400
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Caption         =   "I   n   t   r   o   d   u   c   t   i   o   n"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000015&
            Height          =   255
            Left            =   300
            TabIndex        =   42
            Top             =   240
            Width           =   5400
         End
      End
      Begin VB.Frame Frame13 
         Height          =   5550
         Left            =   -74820
         TabIndex        =   37
         Top             =   840
         Width           =   6000
         Begin MSComDlg.CommonDialog cdlShow 
            Left            =   3600
            Top             =   2760
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox txtOpp 
            Height          =   315
            Left            =   1050
            ScrollBars      =   2  'Vertical
            TabIndex        =   75
            Top             =   1920
            Width           =   1800
         End
         Begin VB.TextBox txtAnc 
            Height          =   1080
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   74
            Top             =   480
            Width           =   4650
         End
         Begin VB.PictureBox picOpp 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   1350
            Left            =   4380
            MouseIcon       =   "frmAdd.frx":08D0
            MousePointer    =   99  'Custom
            ScaleHeight     =   88
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   88
            TabIndex        =   38
            Top             =   1920
            Width           =   1350
         End
         Begin VB.Label lblPath 
            Height          =   450
            Left            =   1050
            TabIndex        =   78
            Top             =   3330
            Width           =   4650
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Caption         =   "O   p   p   o   n   e   n   t"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000015&
            Height          =   255
            Left            =   1050
            TabIndex        =   77
            Top             =   1680
            Width           =   4650
         End
         Begin VB.Label Label6 
            Caption         =   "Name"
            Height          =   210
            Left            =   180
            TabIndex        =   76
            Top             =   1980
            Width           =   810
         End
         Begin VB.Label Label29 
            Caption         =   "Show"
            Height          =   210
            Left            =   180
            TabIndex        =   71
            Top             =   540
            Width           =   810
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Caption         =   "I   n   t   r   o   d   u   c   e"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000015&
            Height          =   255
            Left            =   1050
            TabIndex        =   39
            Top             =   240
            Width           =   4650
         End
      End
      Begin VB.CheckBox chkBat 
         Alignment       =   1  'Right Justify
         Caption         =   "Enable"
         Height          =   210
         Left            =   -74760
         TabIndex        =   36
         Top             =   540
         Width           =   1125
      End
      Begin VB.Frame Frame18 
         Height          =   5550
         Left            =   -74820
         TabIndex        =   34
         Top             =   840
         Width           =   6000
         Begin VB.TextBox txtWin 
            Height          =   4800
            Left            =   300
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   35
            Top             =   480
            Width           =   5400
         End
         Begin VB.Label Label43 
            Alignment       =   2  'Center
            Caption         =   "W   i   n   n   i   n   g"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000015&
            Height          =   210
            Left            =   300
            TabIndex        =   61
            Top             =   240
            Width           =   5400
         End
      End
      Begin VB.Frame Frame8 
         Height          =   5550
         Left            =   -74820
         TabIndex        =   20
         Top             =   840
         Width           =   6000
         Begin VB.ComboBox cmbDia 
            Height          =   315
            ItemData        =   "frmAdd.frx":0BDA
            Left            =   1050
            List            =   "frmAdd.frx":0BE1
            Style           =   2  'Dropdown List
            TabIndex        =   83
            Top             =   1920
            Width           =   1800
         End
         Begin VB.ListBox lstDia 
            Height          =   2595
            Index           =   1
            ItemData        =   "frmAdd.frx":0BEF
            Left            =   2880
            List            =   "frmAdd.frx":0BF1
            TabIndex        =   32
            Top             =   2730
            Width           =   2775
         End
         Begin VB.ListBox lstDia 
            Height          =   2595
            Index           =   0
            ItemData        =   "frmAdd.frx":0BF3
            Left            =   1050
            List            =   "frmAdd.frx":0BF5
            TabIndex        =   31
            Top             =   2730
            Width           =   1800
         End
         Begin VB.CommandButton cmdX 
            Caption         =   "x"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   660
            TabIndex        =   30
            Top             =   5010
            Width           =   315
         End
         Begin VB.CommandButton cmdV 
            Caption         =   "v"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   5400
            TabIndex        =   29
            Top             =   2325
            Width           =   300
         End
         Begin VB.TextBox txtDia 
            Height          =   315
            Left            =   1050
            TabIndex        =   28
            Top             =   2325
            Width           =   4260
         End
         Begin VB.TextBox txtDesC 
            Height          =   1080
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   21
            Top             =   480
            Width           =   4650
         End
         Begin VB.Label Label4 
            Caption         =   "Reply"
            Height          =   210
            Left            =   180
            TabIndex        =   64
            Top             =   2385
            Width           =   900
         End
         Begin VB.Label Label7 
            Caption         =   "Show"
            Height          =   210
            Left            =   180
            TabIndex        =   25
            Top             =   540
            Width           =   900
         End
         Begin VB.Label Label5 
            Caption         =   "Subject"
            Height          =   210
            Left            =   180
            TabIndex        =   24
            Top             =   1980
            Width           =   900
         End
         Begin VB.Label Label50 
            Alignment       =   2  'Center
            Caption         =   "D   e   s   c   r   i   p   t   i   o   n"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000015&
            Height          =   255
            Left            =   1050
            TabIndex        =   23
            Top             =   240
            Width           =   4650
         End
         Begin VB.Label Label51 
            Alignment       =   2  'Center
            Caption         =   "D   i   a   l   o   g"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000015&
            Height          =   255
            Left            =   1050
            TabIndex        =   22
            Top             =   1680
            Width           =   4650
         End
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   0
         Top             =   480
         Width           =   1800
      End
      Begin VB.Frame Frame1 
         Height          =   5550
         Left            =   180
         TabIndex        =   9
         Top             =   840
         Width           =   6000
         Begin VB.CheckBox chkPro 
            Alignment       =   1  'Right Justify
            Caption         =   "Wearable"
            Height          =   210
            Index           =   2
            Left            =   150
            TabIndex        =   82
            Top             =   3120
            Width           =   1080
         End
         Begin VB.OptionButton optOpn 
            Alignment       =   1  'Right Justify
            Caption         =   "Closed"
            Enabled         =   0   'False
            Height          =   210
            Index           =   1
            Left            =   2970
            TabIndex        =   80
            Top             =   2400
            Width           =   1125
         End
         Begin VB.OptionButton optOpn 
            Alignment       =   1  'Right Justify
            Caption         =   "Opened"
            Enabled         =   0   'False
            Height          =   210
            Index           =   0
            Left            =   1500
            TabIndex        =   79
            Top             =   2400
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.TextBox txtDes 
            Height          =   1080
            Left            =   1050
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   480
            Width           =   4650
         End
         Begin VB.ComboBox cmbCon 
            Height          =   315
            ItemData        =   "frmAdd.frx":0BF7
            Left            =   3900
            List            =   "frmAdd.frx":0BFE
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1920
            Width           =   1800
         End
         Begin VB.ComboBox cmbSize 
            Height          =   315
            ItemData        =   "frmAdd.frx":0C0C
            Left            =   1050
            List            =   "frmAdd.frx":0C1F
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1920
            Width           =   1800
         End
         Begin VB.CheckBox chkPro 
            Alignment       =   1  'Right Justify
            Caption         =   "Openable"
            Height          =   210
            Index           =   0
            Left            =   150
            TabIndex        =   11
            Top             =   2400
            Width           =   1080
         End
         Begin VB.CheckBox chkPro 
            Alignment       =   1  'Right Justify
            Caption         =   "Stationary"
            Height          =   210
            Index           =   1
            Left            =   150
            TabIndex        =   10
            Top             =   2760
            Width           =   1080
         End
         Begin VB.Label lblCon 
            Height          =   210
            Left            =   210
            TabIndex        =   33
            Top             =   3820
            Width           =   750
         End
         Begin VB.Label Label2 
            Caption         =   "Show"
            Height          =   210
            Left            =   180
            TabIndex        =   19
            Top             =   540
            Width           =   810
         End
         Begin VB.Label Label22 
            Caption         =   "Size"
            Height          =   210
            Left            =   180
            TabIndex        =   18
            Top             =   1980
            Width           =   810
         End
         Begin VB.Label Label26 
            Caption         =   "Inside"
            Height          =   210
            Left            =   3000
            TabIndex        =   17
            Top             =   1980
            Width           =   810
         End
         Begin VB.Label Label47 
            Alignment       =   2  'Center
            Caption         =   "P   r   o   p   e   r   t   i   e   s"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000015&
            Height          =   255
            Left            =   1050
            TabIndex        =   16
            Top             =   1680
            Width           =   4650
         End
         Begin VB.Label Label46 
            Alignment       =   2  'Center
            Caption         =   "D   e   s   c   r   i   p   t   i   o   n"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000015&
            Height          =   255
            Left            =   1050
            TabIndex        =   15
            Top             =   240
            Width           =   4650
         End
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Index           =   1
         Left            =   -73800
         TabIndex        =   1
         Top             =   480
         Width           =   1800
      End
      Begin VB.Label Label10 
         Caption         =   "After Task"
         Height          =   210
         Left            =   -71850
         TabIndex        =   73
         Top             =   540
         Width           =   810
      End
      Begin VB.Label Label28 
         Caption         =   "Show after"
         Height          =   210
         Left            =   -72000
         TabIndex        =   70
         Top             =   6660
         Width           =   900
      End
      Begin VB.Label Label15 
         Caption         =   "Show after"
         Height          =   210
         Left            =   -72000
         TabIndex        =   68
         Top             =   6660
         Width           =   900
      End
      Begin VB.Label Label55 
         Caption         =   "Show after"
         Height          =   210
         Left            =   3000
         TabIndex        =   66
         Top             =   6660
         Width           =   900
      End
      Begin VB.Label Label9 
         Caption         =   "* Action"
         Height          =   210
         Left            =   -71850
         TabIndex        =   59
         Top             =   540
         Width           =   810
      End
      Begin VB.Label Label62 
         Caption         =   "* Name"
         Height          =   210
         Left            =   -74760
         TabIndex        =   58
         Top             =   540
         Width           =   750
      End
      Begin VB.Label Label12 
         Caption         =   "Subject"
         Height          =   210
         Left            =   -74655
         TabIndex        =   52
         Top             =   540
         Width           =   810
      End
      Begin VB.Label Label40 
         Caption         =   "Start in Room"
         Height          =   210
         Left            =   -72000
         TabIndex        =   47
         Top             =   6660
         Width           =   1200
      End
      Begin VB.Label Label24 
         Caption         =   "Author"
         Height          =   210
         Left            =   -74655
         TabIndex        =   46
         Top             =   960
         Width           =   750
      End
      Begin VB.Label Label23 
         Caption         =   "Title"
         Height          =   210
         Left            =   -74655
         TabIndex        =   45
         Top             =   540
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "* Name"
         Height          =   210
         Left            =   240
         TabIndex        =   27
         Top             =   540
         Width           =   750
      End
      Begin VB.Label Label8 
         Caption         =   "* Name"
         Height          =   210
         Left            =   -74760
         TabIndex        =   26
         Top             =   540
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strElm As String, iTab As Integer

Private Sub Form_Load()
  frmMaker.Enabled = False
End Sub

Public Sub LoadTab(iNum As Integer, strName As String)
  Dim strPro As String, strAct As String, strTmp As String
  iTab = iNum: strElm = strName
  
  For i% = 0 To 6
    SSTab.TabVisible(i%) = IIf(i% = iNum, True, False)
  Next i%
  
  strPro = ReadINI("Properties", strName, strFile)
  strAct = ReadINI("Actions", strName, strFile)

  Select Case iNum
    Case 0
      txtName(iNum) = strName
      ObjToObj cmbCon, frmMaker.lstTmp(1)
      For i% = 0 To frmMaker.lstTmp(0).ListCount - 1
        strTmp = ReadINI("Properties", frmMaker.lstTmp(0).List(i%), strFile)
        If PStr("Find", strTmp, , "openable") > 0 And frmMaker.lstTmp(0).List(i%) <> strName Then cmbCon.AddItem frmMaker.lstTmp(0).List(i%)
      Next i%
      frmMaker.ReloadList cmbAfter(iNum), -1, "<Nothing>"
      txtDes = PStr("Find", strAct, "Return", "look")
      If PStr("Find", strPro, , "openable") > 0 Then chkPro(0).Value = 1
      If PStr("Find", strPro, , "stationary") > 0 Then chkPro(1).Value = 1
      If PStr("Find", strPro, , "wearable") > 0 Then chkPro(2).Value = 1
      If PStr("Find", strPro, "Return", "openable") = "closed" Then optOpn(1).Value = True
      cmbCon.ListIndex = GetIndex(cmbCon, PStr("Find", strPro, "Return", "container"))
      cmbSize.ListIndex = IIf(strName = "", 2, PStr("Find", strPro, "Return", "size"))
      cmbAfter(iNum).ListIndex = GetIndex(cmbAfter(iNum), PStr("Find", strPro, "Return", "after"))

    Case 1
      txtName(iNum) = strName
      ObjToObj cmbDia, frmMaker.lstTmp(0)
      ObjToObj cmbDia, frmMaker.lstTmp(1)
      ObjToObj cmbDia, frmMaker.lstTmp(2)
      cmbDia.ListIndex = 0
      frmMaker.ReloadList cmbAfter(iNum), -1, "<Nothing>"
      txtDesC = PStr("Find", strAct, "Return", "look")
      varTmp = Split(ReadINI("Actions", strName & "/ask", strFile), "|")
      For i% = 1 To UBound(varTmp)
        lstDia(0).AddItem Split(varTmp(i%), "")(0): lstDia(1).AddItem Split(varTmp(i%), "")(1)
      Next i%
      cmbAfter(iNum).ListIndex = GetIndex(cmbAfter(iNum), PStr("Find", strPro, "Return", "after"))
   
    Case 2
      txtName(iNum) = strName
      frmMaker.ReloadList cmbAfter(iNum), -1, "<Nothing>"
      txtDesO = PStr("Find", strAct, "Return", "look")
      If strName <> "" Then If PStr("Find", strPro, , "stationary") = 0 Then chkDin.Value = 1
      cmbAfter(iNum).ListIndex = GetIndex(cmbAfter(iNum), PStr("Find", strPro, "Return", "after"))
      If strName <> "" Then If PStr("Find", strPro, "Return", "after") = strName Then chkInv.Value = 1
    
    Case 3
      If strName <> "" Then txtSub = Split(strName, "/")(0): txtAct = Split(strName, "/")(1)
      varTmp = Split(ReadINI("Element", 0, strFile), "|")
      For i% = 1 To UBound(varTmp)
        If Split(varTmp(i%), "")(0) <> LCase(txtSub) Then cmbUse.AddItem Split(varTmp(i%), "")(0)
      Next i%
      For i% = 0 To frmMaker.lstTmp(0).ListCount - 1
        strTmp = ReadINI("Properties", frmMaker.lstTmp(0).List(i%), strFile)
        If PStr("Find", strTmp, "Return", "container") = txtSub Then cmbDrop.AddItem frmMaker.lstTmp(0).List(i%)
      Next i%
      If cmbDrop.ListCount > 2 Then cmbDrop.AddItem "<All>"
      txtNdes = PStr("Find", strPro, "Return", "Ndes")
      txtMsg = PStr("Find", ReadINI("Actions", txtSub, strFile), "Return", txtAct)
      If PStr("Find", strPro, , "hide") > 0 Then chkHide.Value = 1
      If PStr("Find", strAct, , "permanent") > 0 Then chkPer.Value = 1
      If PStr("Find", strPro, , "dice") > 0 Then chkDice.Value = 1
      If txtAct = "" Then If PStr("Find", ReadINI("Properties", txtSub, strFile), , "stationary") > 0 Then optSta.Value = True
      If txtAct <> "" Then If PStr("Find", strPro, "Return", "stationary") = "yes" Then optSta.Value = True
      cmbUse.ListIndex = GetIndex(cmbUse, PStr("Find", strPro, "Return", "use"))
      cmbDrop.ListIndex = GetIndex(cmbDrop, PStr("Find", strPro, "Return", "drop"))
    
    Case 4
      varTmp = Split(ReadINI("Starting", "Info", strFile), "|")
      If UBound(varTmp) > -1 Then
        txtTitle = varTmp(1): txtAuthor = varTmp(2): txtIntro = varTmp(4)
        For i% = 5 To UBound(varTmp): txtIntro = txtIntro & vbCrLf & varTmp(i%): Next i%
      End If
      varTmp = Split(ReadINI("Map", "Title", strFile), "|")
      For i% = 1 To UBound(varTmp): cmbStart.AddItem Split(varTmp(i%), "")(1): Next i%
      cmbStart.ListIndex = PStr("Find", ReadINI("Map", "Title", strFile), , ReadINI("Starting", "Start", strFile)) - 1

    Case 5
      strTmp = ReadINI("Ending", "Battle", strFile)
      If PStr("Find", strTmp, , "Enabled") > 0 Then chkBat.Value = 1
      txtAnc = PStr("Find", strTmp, "Return", "Introduce")
      txtOpp = PStr("Find", strTmp, "Return", "Opponent")
      lblPath = PStr("Find", strTmp, "Return", "Picture")
      picOpp.Picture = LoadPicture(lblPath)
      
    Case 6
      frmMaker.ReloadList cmbAfter(iNum), -1, "<None>"
      txtWin = Replace(ReadINI("Ending", "Win", strFile), "|", vbCrLf)
      strTmp = ReadINI("Ending", "Final", strFile)
      GetIndex cmbAfter(iNum), strTmp, "Remove"
      If strTmp <> "" Then cmbAfter(iNum).AddItem strTmp
      cmbAfter(iNum).ListIndex = GetIndex(cmbAfter(iNum), strTmp)
  End Select
  frmAdd.Show
End Sub

Public Sub SaveTab(iNum As Integer, strName As String)
  Dim strPro As String, strAct As String, strTmp As String, strHeld As String
  
  strAct = ReadINI("Actions", strName, strFile)
  strPro = ReadINI("Properties", strName, strFile)
  strHeld = ReadINI("Properties", cmbCon.Text, strFile)
  
  Select Case iNum
    Case 0
      If SetName(strName) = True Then Exit Sub
      For i% = 0 To frmMaker.lstTmp(0).ListCount - 1
        strTmp = ReadINI("Properties", frmMaker.lstTmp(0).List(i%), strFile)
        If PStr("Find", strTmp, "Return", "index") > PStr("Find", strPro, "Return", "index") And PStr("Find", strTmp, "Return", "parent") = PStr("Find", strPro, "Return", "parent") Then
          If cmbCon.ListIndex > 0 Then If PStr("Find", strHeld, "Return", "parent") <> PStr("Find", strPro, "Return", "parent") Then strTmp = PStr("Add", strTmp, PStr("Find", strHeld, "Return", "parent"), "parent")
          If cmbCon.ListIndex > 0 Then If PStr("Find", strHeld, "Return", "index") = PStr("Find", strPro, "Return", "index") Then strTmp = PStr("Add", strTmp, PStr("Find", strTmp, "Return", "index") + 1, "index")
          If chkPro(0).Value = 1 And optOpn(0).Value = True Then If PStr("Find", strTmp, , "after") = 0 Or Right(PStr("Find", strTmp, "Return", "after"), 5) = "/open" Then If Right(PStr("Find", strHeld, "Return", "after"), 5) = "/open" Then strTmp = PStr("Add", strTmp, PStr("Find", strHeld, "Return", "after"), "after") Else strTmp = PStr("Remove", strTmp, "", "after")
          If chkPro(0).Value = 1 And optOpn(1).Value = True Then If PStr("Find", strTmp, , "after") = 0 Or Right(PStr("Find", strTmp, "Return", "after"), 5) = "/open" Then strTmp = PStr("Add", strTmp, strName & "/open", "after")
        End If
        WriteINI "Properties", frmMaker.lstTmp(0).List(i%), strTmp, strFile
      Next i%
      strPro = PStr(IIf(cmbAfter(iNum).ListIndex > 0, "Add", "Remove"), strPro, cmbAfter(iNum).Text, "after") '1
      If Right(PStr("Find", strHeld, "Return", "after"), 5) = "/open" Then strPro = PStr("Add", strPro, PStr("Find", strHeld, "Return", "after"), "after") '2
      If PStr("Find", strHeld, "Return", "openable") = "closed" Then strPro = PStr("Add", strPro, cmbCon.Text & "/open", "after") '3
      strPro = PStr(IIf(chkPro(0).Value = 1, "Add", "Remove"), strPro, IIf(optOpn(0).Value = True, "opened", "closed"), "openable")
      strPro = PStr("Add", strPro, IIf(cmbCon.ListIndex = 0, strName, PStr("Find", strHeld, "Return", "parent")), "parent")
      strPro = PStr("Add", strPro, IIf(cmbCon.ListIndex = 0, -1, PStr("Find", strHeld, "Return", "index")) + 1, "index")
      strPro = PStr(IIf(cmbCon.ListIndex > 0, "Add", "Remove"), strPro, cmbCon.Text, "container")
      strPro = PStr(IIf(chkPro(1).Value = 1, "Add", "Remove"), strPro, "", "stationary")
      strPro = PStr(IIf(chkPro(2).Value = 1, "Add", "Remove"), strPro, "", "wearable")
      strAct = PStr(IIf(txtDes <> "", "Add", "Remove"), strAct, txtDes, "look")
      strPro = PStr("Add", strPro, cmbSize.ListIndex, "size")
      strPro = PStr("Add", strPro, "object", "element")
      WriteINI "Properties", strName, strPro, strFile
      WriteINI "Actions", strName, strAct, strFile

    Case 1
      If SetName(strName) = True Then Exit Sub
      For i% = 0 To lstDia(0).ListCount - 1
        strTmp = PStr("Add", strTmp, lstDia(1).List(i%), lstDia(0).List(i%))
        WriteINI "Actions", strName & "/ask", strTmp, strFile
      Next i%
      strAct = PStr(IIf(txtDesC <> "", "Add", "Remove"), strAct, txtDesC, "look")
      strPro = PStr(IIf(cmbAfter(iNum).ListIndex > 0, "Add", "Remove"), strPro, cmbAfter(iNum).Text, "after")
      strPro = PStr("Add", strPro, "character", "element")
      strPro = PStr("Add", strPro, "closed", "openable")
      strPro = PStr("Add", strPro, strName, "parent")
      strPro = PStr("Add", strPro, "", "stationary")
      strPro = PStr("Add", strPro, 0, "index")
      WriteINI "Properties", strName, strPro, strFile
      WriteINI "Actions", strName, strAct, strFile
    
    Case 2
      If SetName(strName) = True Then Exit Sub
      strAct = PStr(IIf(txtDesO <> "", "Add", "Remove"), strAct, txtDesO, "look")
      If chkInv.Value = 1 Then strPro = PStr("Add", strPro, strName, "after") Else strPro = PStr(IIf(cmbAfter(iNum).ListIndex > 0, "Add", "Remove"), strPro, cmbAfter(iNum).Text, "after")
      strPro = PStr(IIf(chkDin.Value = 0, "Add", "Remove"), strPro, "", "stationary")
      strPro = PStr("Add", strPro, "other", "element")
      WriteINI "Properties", strName, strPro, strFile
      WriteINI "Actions", strName, strAct, strFile
    
    Case 3
      If strName <> "" Then txtSub = Split(strName, "/")(0): txtAct = LCase(Split(strName, "/")(1))
      If txtAct = "" Then MsgBox "Information required (*)", vbInformation: Exit Sub
      strPro = PStr(IIf(cmbUse.ListIndex > 0, "Add", "Remove"), strPro, cmbUse.Text, "use")
      strPro = PStr(IIf(cmbDrop.ListIndex > 0, "Add", "Remove"), strPro, cmbDrop.Text, "drop")
      strPro = PStr("Add", strPro, IIf(optSta.Value = True, "yes", "no"), "stationary")
      strAct = PStr(IIf(chkPer.Value = 1, "Add", "Remove"), strAct, , "permanent")
      strPro = PStr(IIf(txtNdes <> "", "Add", "Remove"), strPro, txtNdes, "Ndes")
      strPro = PStr(IIf(chkDice.Value = 1, "Add", "Remove"), strPro, , "dice")
      strPro = PStr(IIf(chkHide.Value = 1, "Add", "Remove"), strPro, , "hide")
      WriteINI "Properties", strName, strPro, strFile
      WriteINI "Actions", strName, strAct, strFile
      WriteINI "Actions", txtSub, PStr("Add", ReadINI("Actions", txtSub, strFile), txtMsg, txtAct), strFile

    Case 4
      If txtTitle = "" Then txtTitle = "Untitled"
      If txtAuthor = "" Then txtAuthor = "Anonymous"
      WriteINI "Starting", "Info", "|" & txtTitle & "|" & txtAuthor & "||" & Replace(txtIntro, vbCrLf, "|"), strFile
      WriteINI "Starting", "Start", PStr("Find", ReadINI("Map", "Title", strFile), , "Return", cmbStart.ListIndex + 1), strFile

    Case 5
      strTmp = PStr("Add", strTmp, txtOpp, "Opponent")
      strTmp = PStr(IIf(lblPath <> "", "Add", "Remove"), strTmp, lblPath, "Picture")
      strTmp = PStr("Add", strTmp, txtAnc, "Introduce")
      strTmp = PStr(IIf(chkBat.Value = 1, "Add", "Remove"), strTmp, , "Enabled")
      WriteINI "Ending", "Battle", strTmp, strFile
    
    Case 6
      WriteINI "Ending", "Final", cmbAfter(iNum).Text, strFile
      WriteINI "Ending", "Win", Replace(txtWin, vbCrLf, "|"), strFile
  End Select
  strElm = IIf(iNum < 3, txtName(iNum), "")
  Unload Me
End Sub

Private Sub picOpp_Click()
On Error Resume Next
  cdlShow.Filter = "All Picture Files|*.bmp;*.dib;*.gif;*.jpg;*.wmf;*.emf;*.ico"
  cdlShow.InitDir = App.Path
  cdlShow.filename = ""
  cdlShow.ShowOpen
  lblPath = cdlShow.filename
  If Err = 0 Then picOpp.Picture = LoadPicture(lblPath)
End Sub

Private Sub chkHide_Click()
  optSta.Enabled = IIf(chkHide.Value = 0, True, False): optDin.Enabled = optSta.Enabled: txtNdes.Enabled = optSta.Enabled: chkPer.Enabled = optSta.Enabled
End Sub

Private Function SetName(strName As String) As Boolean
  If PStr("Find", ReadINI("Element", iTab, strFile), , LCase(strName)) > 0 Then MsgBox "Repeated Name", vbInformation: SetName = True
  If strName = "" Then MsgBox "Information required (*)", vbInformation: SetName = True
End Function

Private Sub chkPro_Click(Index As Integer)
  optOpn(0).Enabled = chkPro(0).Value: optOpn(1).Enabled = chkPro(0).Value
End Sub

Private Sub chkInv_Click()
  cmbAfter(2).Enabled = IIf(chkInv.Value = 0, True, False): cmbAfter(2).ListIndex = 0: txtDesO.Enabled = cmbAfter(2).Enabled: chkDin.Enabled = cmbAfter(2).Enabled
End Sub

Private Sub lstDia_DblClick(Index As Integer)
  cmbDia.ListIndex = GetIndex(cmbDia, lstDia(0).Text)
  txtDia = lstDia(1).Text
End Sub

Private Sub lstDia_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  lstDia(IIf(Index = 0, 1, 0)).ListIndex = lstDia(Index).ListIndex: cmdX.Enabled = True
End Sub

Private Sub cmdV_Click()
  If cmbDia.ListIndex > 0 And txtDia <> "" Then lstDia(0).AddItem cmbDia.Text: lstDia(1).AddItem txtDia: txtDia = "": cmbDia.ListIndex = 0
End Sub

Private Sub cmdX_Click()
  lstDia(0).RemoveItem lstDia(0).ListIndex: lstDia(1).RemoveItem lstDia(1).ListIndex: cmdX.Enabled = False
End Sub

Private Sub cmdAdd_Click()
  SaveTab iTab, IIf(iTab < 3, txtName(iTab), txtSub & "/" & LCase(txtAct))
End Sub

Private Sub cmdCnl_Click()
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmMaker.AddElement strElm, iTab
  frmMaker.Enabled = True
End Sub
