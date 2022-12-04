VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmUserTypes 
   BackColor       =   &H00FF8080&
   Caption         =   "User Types"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4020
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserTypes.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   4020
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF8080&
      Caption         =   "Employee User"
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
      Index           =   9
      Left            =   600
      TabIndex        =   10
      Top             =   5040
      Width           =   2895
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   1095
      Left            =   2400
      Picture         =   "frmUserTypes.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   1095
      Left            =   600
      Picture         =   "frmUserTypes.frx":128E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6120
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF8080&
      Caption         =   "Pharmacy User"
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
      Index           =   7
      Left            =   600
      TabIndex        =   7
      Top             =   4560
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF8080&
      Caption         =   "Patient User"
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
      Index           =   6
      Left            =   600
      TabIndex        =   6
      Top             =   4080
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF8080&
      Caption         =   "Hospital Manager"
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
      Index           =   5
      Left            =   600
      TabIndex        =   5
      Top             =   3600
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF8080&
      Caption         =   "Employee Manager"
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
      Index           =   4
      Left            =   600
      TabIndex        =   4
      Top             =   3120
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF8080&
      Caption         =   "Pharmacy Manager"
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
      Index           =   3
      Left            =   600
      TabIndex        =   3
      Top             =   2640
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF8080&
      Caption         =   "Patient Manager"
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
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   2160
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF8080&
      Caption         =   "Guest"
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
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FF8080&
      Caption         =   "Administrator"
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
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Value           =   -1  'True
      Width           =   2895
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      FormHeightDT    =   8400
      FormWidthDT     =   4140
      FormScaleHeightDT=   7890
      FormScaleWidthDT=   4020
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "USER TYPES"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   960
      TabIndex        =   11
      Top             =   240
      Width           =   2460
   End
   Begin VB.Shape Shape1 
      Height          =   4575
      Left            =   360
      Top             =   1080
      Width           =   3375
   End
End
Attribute VB_Name = "frmUserTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
frmUserDetails.Enabled = True
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim sel As String

If Option1(0).Value = True Then ' if admin
    sel = "1"
ElseIf Option1(1).Value = True Then 'if Guest
    sel = "0"
ElseIf Option1(2).Value = True Then 'if Patient Manager
    sel = "3"
ElseIf Option1(3).Value = True Then 'if Pharmacy Manager
    sel = "4"
ElseIf Option1(4).Value = True Then 'if Employee Manager
    sel = "2"
ElseIf Option1(5).Value = True Then 'if Hospital Manager
    sel = "5"
ElseIf Option1(6).Value = True Then 'if Patient User
    sel = "7"
ElseIf Option1(7).Value = True Then 'if pharmacy User
    sel = "8"
ElseIf Option1(8).Value = True Then 'if Employee User
    sel = "6"
End If

frmUserDetails.Enabled = True
frmUserDetails.txtFields(2) = sel

Me.Hide



End Sub

Private Sub Form_Activate()
frmUserDetails.Enabled = False
End Sub

Private Sub Form_Deactivate()
frmUserDetails.Enabled = True
End Sub

Private Sub Form_Load()
frmUserDetails.Enabled = False
End Sub
