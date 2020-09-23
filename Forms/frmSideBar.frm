VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5734474E-78D3-4254-99B9-C35F31BDF509}#63.0#0"; "vbskpro2.ocx"
Begin VB.Form frmSideBar 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Side Bar"
   ClientHeight    =   8145
   ClientLeft      =   -135
   ClientTop       =   435
   ClientWidth     =   2895
   Icon            =   "frmSideBar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   2895
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Height          =   3855
      Left            =   360
      ScaleHeight     =   3795
      ScaleWidth      =   2115
      TabIndex        =   12
      Top             =   480
      Width           =   2175
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   240
         TabIndex        =   18
         Top             =   3120
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         Caption         =   "&Log Off"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   240
         MaskColor       =   &H00FFFF80&
         TabIndex        =   17
         Top             =   2520
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         Caption         =   "Disable Skin"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   240
         TabIndex        =   16
         Top             =   1965
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Caption         =   "Change Password"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   240
         TabIndex        =   15
         Top             =   1395
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Caption         =   "&Calculator"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   240
         TabIndex        =   14
         Top             =   810
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "&Note Pad"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   355
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "       User Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   4800
      Width           =   2655
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   2535
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   100
         Picture         =   "frmSideBar.frx":57E2
         Top             =   -30
         Width           =   240
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "      Time Log-in"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   5760
      Width           =   2655
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   2535
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   100
         Picture         =   "frmSideBar.frx":5D6C
         Top             =   -30
         Width           =   240
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "       Menu Explorer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4515
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2655
      Begin VB.PictureBox Picture1 
         Height          =   3660
         Left            =   240
         ScaleHeight     =   3600
         ScaleWidth      =   2115
         TabIndex        =   6
         Top             =   360
         Width           =   2175
         Begin MSComctlLib.TreeView TreeView1 
            Height          =   3600
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   6350
            _Version        =   393217
            HideSelection   =   0   'False
            Indentation     =   529
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            FullRowSelect   =   -1  'True
            SingleSel       =   -1  'True
            ImageList       =   "ImageList1"
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   120
         Picture         =   "frmSideBar.frx":60F6
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "      TODAY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   6600
      Width           =   2655
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   2040
         Top             =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "11/23/03"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label3 
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "04:12 AM"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
      Begin VB.Image Image4 
         Height          =   240
         Left            =   120
         Picture         =   "frmSideBar.frx":6480
         Top             =   0
         Width           =   240
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSideBar.frx":680A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSideBar.frx":6BA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSideBar.frx":6F48
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSideBar.frx":72EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSideBar.frx":7690
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSideBar.frx":7A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSideBar.frx":7DD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSideBar.frx":8174
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSideBar.frx":8510
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin vbskpro.Skinner Skinner1 
      Left            =   2040
      Top             =   5280
      _ExtentX        =   1270
      _ExtentY        =   1270
      OldForeColor    =   0
      SysDisableSkinCaption=   "&Disable Skin"
      LcK1            =   "3.66*/4/0*/1-5*210/."
      LcK2            =   $"frmSideBar.frx":88AC
      AmbientB        =   ";<=>?7B:><7=<A<7CC;@"
   End
End
Attribute VB_Name = "frmSideBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Shell "Notepad.exe", vbNormalFocus
End Sub

Private Sub Command2_Click()
Shell "calc.exe", vbNormalFocus
End Sub

Private Sub Command3_Click()
MDIMain.chngPass_Click
End Sub

Private Sub Command4_Click()
    Skinner1.Enabled = Not Skinner1.Enabled
    If Skinner1.Enabled Then
        Command4.Caption = "Disable &skin"
    Else
        Command4.Caption = "Enable &skin"
    End If
End Sub

Private Sub Command5_Click()
MDIMain.logoff_Click
End Sub

Private Sub Command6_Click()

Unload MDIMain
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0

Text1.Text = "Logged in as : " & User
Text2.Text = LogTime
Label2.Caption = Date

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.sidebar.Checked = False
End Sub

Private Sub Timer1_Timer()
Label4.Caption = Time
End Sub
