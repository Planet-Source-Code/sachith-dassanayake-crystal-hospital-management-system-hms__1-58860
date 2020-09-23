VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplashScreen 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Crytal Hospital Management System"
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   240
   End
   Begin MSComctlLib.ProgressBar ProgLoad 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   3660
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "License To: Crystal Hospital PVT LTD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   4
      Top             =   2640
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Copyright 2005"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   4125
      Left            =   0
      Picture         =   "frmSplashScreen.frx":0000
      Top             =   0
      Width           =   3750
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   6600
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crystal Hospital Management System"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   1695
      Left            =   4080
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmSplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All form ontop stuff :)
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Sub Form_Activate()
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Load()
    
    'Centers the form.
    Left = (Screen.Width - Width) \ 2
    Top = (Screen.Height - Height) \ 2

End Sub

Private Sub Timer1_Timer()
    
    ProgLoad.Value = ProgLoad.Value + 5
    'If the Progress Bar (ProgLoad) is 100% then your function happens.
    If ProgLoad.Value = 100 Then
        
        'Your function, can be anything. Open another form, frmMain.show... Ect.
        frmLogin.Show
        'Unloads this form
        Unload Me
    End If

End Sub
