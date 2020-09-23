VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmCredits 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Us"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6435
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   5400
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      AutoCenterForm  =   -1  'True
      FormHeightDT    =   4530
      FormWidthDT     =   6525
      FormScaleHeightDT=   4050
      FormScaleWidthDT=   6435
   End
   Begin VB.PictureBox P 
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
      Height          =   1575
      Index           =   0
      Left            =   3360
      ScaleHeight     =   1515
      ScaleWidth      =   2835
      TabIndex        =   5
      Top             =   240
      Width           =   2895
      Begin VB.Label L3 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Sachith Dassanayake"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label L4 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "sachithd@gmail.com +9411-2614974"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   6
         Top             =   1200
         Width           =   2895
      End
   End
   Begin VB.PictureBox P 
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
      Height          =   1575
      Index           =   1
      Left            =   120
      ScaleHeight     =   1515
      ScaleWidth      =   2835
      TabIndex        =   1
      Top             =   240
      Width           =   2895
      Begin VB.Label L2 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "nicefunnyguy@hotmail.com +9411-2593741"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   4
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label L1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Asvine Ganeshan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   840
         Width           =   2895
      End
   End
   Begin VB.Timer T 
      Interval        =   100
      Left            =   3360
      Top             =   3120
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Close"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Timer tmrScrollTitle 
      Interval        =   100
      Left            =   3840
      Top             =   3120
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.CrystalHMS.com"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   1320
      MouseIcon       =   "frmCredits.frx":0E42
      TabIndex        =   3
      Top             =   2400
      Width           =   4455
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
MDIMain.Enabled = False
SetInitialCaption "Crystal Hospital Management System", 80, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.Enabled = True

End Sub

Private Sub Label1_Click()
ShellExecute hWnd, "open", "http://crystalhms.freesuperhost.com", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub T_Timer()
If L1.Top <= -800 Then L1.Top = 1560
If L2.Top <= -800 Then L2.Top = 1560
If L3.Top <= -800 Then L1.Top = 1560
If L4.Top <= -800 Then L2.Top = 1560

L1.Top = L1.Top - 15
L2.Top = L2.Top - 15
L3.Top = L1.Top - 15
L4.Top = L2.Top - 15
End Sub

Private Sub tmrScrollTitle_Timer()
    ScrollTitle "Crystal Hospital Management System", 80, Me
End Sub
