VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmAppRegister 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Register Crystal HMS"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAppRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
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
      AutoCenterForm  =   -1  'True
      FormHeightDT    =   3675
      FormWidthDT     =   6120
      FormScaleHeightDT=   3195
      FormScaleWidthDT=   6030
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   1
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      MaxLength       =   4
      TabIndex        =   0
      Top             =   1560
      Width           =   735
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Enter your 20 digits registration key provided with the original CD"
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   600
      TabIndex        =   7
      Top             =   360
      Width           =   4935
   End
End
Attribute VB_Name = "frmAppRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()

Unload Me
End Sub

Private Sub Form_Load()
MDIMain.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.Enabled = True
End Sub

Private Sub OKButton_Click()
If Text1 = "Gx1b" And Text2 = "CqLn" And Text3 = "NIPI" And Text4 = "VUOY" And Text5 = "KCuF" Then
    SaveSetting App.Title, "Settings", "CHECK", "ALLOW"
    MsgBox "Thank you for registering Crystal Hospital Management System" & vbCrLf & "Please restart the program for verify the activation key", vbExclamation
    
    Unload Me
Else
    MsgBox "Invalid Registration key", vbInformation, "Invalid Serial Number"
    
    Text5.SetFocus
    
End If


End Sub

Private Sub Text1_Change()
If Len(Text1) = 4 Then
    Text2.SetFocus
End If
End Sub

Private Sub Text2_Change()
If Len(Text2) = 4 Then
    Text3.SetFocus
End If
End Sub

Private Sub Text3_Change()
If Len(Text3) = 4 Then
    Text4.SetFocus
End If
End Sub

Private Sub Text4_Change()
If Len(Text4) = 4 Then
    Text5.SetFocus
End If
End Sub
