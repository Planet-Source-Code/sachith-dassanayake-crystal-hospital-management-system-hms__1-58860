VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmAppPreferences 
   BackColor       =   &H00FF8080&
   Caption         =   "Application Preferences"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAppPreferences.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2205
   ScaleWidth      =   3855
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   480
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      AutoCenterForm  =   -1  'True
      FormHeightDT    =   2715
      FormWidthDT     =   3975
      FormScaleHeightDT=   2205
      FormScaleWidthDT=   3855
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF8080&
      Caption         =   "Show Tips at StartUp"
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
End
Attribute VB_Name = "frmAppPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdsave_click()
If Check1.Value = 1 Then
    SaveSetting App.Title, "Options", "Show Tips at Startup", 1
    
End If


If Check1.Value = 0 Then
     SaveSetting App.Title, "Options", "Show Tips at Startup", 0
    
End If
    Unload Me

End Sub
