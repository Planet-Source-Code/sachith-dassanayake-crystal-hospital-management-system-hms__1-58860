VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmCompany 
   Caption         =   "Hospital Information"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6750
   Icon            =   "frmCompany.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4860
   ScaleWidth      =   6750
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Hospital Information"
      TabPicture(0)   =   "frmCompany.frx":57E2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ActiveResize1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "picComp"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Other Information"
      TabPicture(1)   =   "frmCompany.frx":57FE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      Begin VB.PictureBox picComp 
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   1200
         ScaleHeight     =   3615
         ScaleWidth      =   4905
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   720
         Width           =   4905
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1140
            MaxLength       =   50
            TabIndex        =   8
            Text            =   "Crystal Hospital PVT LTD"
            Top             =   75
            Width           =   3315
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   1
            Left            =   1140
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Text            =   "frmCompany.frx":581A
            Top             =   450
            Width           =   3315
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   1140
            MaxLength       =   50
            TabIndex        =   6
            Text            =   "Sri Lanka"
            Top             =   1350
            Width           =   3315
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   1140
            MaxLength       =   25
            TabIndex        =   5
            Text            =   "0112-593327"
            Top             =   1725
            Width           =   3315
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   1140
            MaxLength       =   25
            TabIndex        =   4
            Text            =   "0112-619957"
            Top             =   2100
            Width           =   3315
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   1140
            MaxLength       =   50
            TabIndex        =   3
            Text            =   "customer@crystalHospital.com"
            Top             =   2475
            Width           =   3315
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   1140
            MaxLength       =   50
            TabIndex        =   2
            Text            =   "http://www.CrystalHospital.com"
            Top             =   2850
            Width           =   3315
         End
         Begin VB.Label Label3 
            Caption         =   "Company:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   75
            TabIndex        =   15
            Top             =   75
            Width           =   885
         End
         Begin VB.Label Label3 
            Caption         =   "Address:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   75
            TabIndex        =   14
            Top             =   450
            Width           =   765
         End
         Begin VB.Label Label3 
            Caption         =   "Country:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   75
            TabIndex        =   13
            Top             =   1350
            Width           =   765
         End
         Begin VB.Label Label3 
            Caption         =   "Phone:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   75
            TabIndex        =   12
            Top             =   1725
            Width           =   765
         End
         Begin VB.Label Label3 
            Caption         =   "Fax:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   75
            TabIndex        =   11
            Top             =   2100
            Width           =   765
         End
         Begin VB.Label Label3 
            Caption         =   "E-Mail:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   5
            Left            =   75
            TabIndex        =   10
            Top             =   2475
            Width           =   765
         End
         Begin VB.Label Label3 
            Caption         =   "Web Site:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   6
            Left            =   75
            TabIndex        =   9
            Top             =   2850
            Width           =   1005
         End
      End
      Begin ActiveResizeCtl.ActiveResize ActiveResize1 
         Left            =   360
         Top             =   1080
         _ExtentX        =   847
         _ExtentY        =   847
         Resolution      =   4
         ScreenHeight    =   1024
         ScreenWidth     =   1280
         ScreenHeightDT  =   1024
         ScreenWidthDT   =   1280
         AutoCenterForm  =   -1  'True
         FormHeightDT    =   5370
         FormWidthDT     =   6870
         FormScaleHeightDT=   4860
         FormScaleWidthDT=   6750
         ResizeFormBackground=   -1  'True
         ResizePictureBoxContents=   -1  'True
      End
   End
End
Attribute VB_Name = "frmCompany"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
