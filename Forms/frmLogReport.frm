VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmLogReport 
   BackColor       =   &H00FF8080&
   Caption         =   "User Log In Report"
   ClientHeight    =   4620
   ClientLeft      =   3240
   ClientTop       =   2565
   ClientWidth     =   8340
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4620
   ScaleWidth      =   8340
   Begin VB.CommandButton cmdViewReport 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Report"
      Height          =   1095
      Left            =   3000
      Picture         =   "frmLogReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   1095
      Left            =   4320
      Picture         =   "frmLogReport.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "DATE RANGE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   7095
      Begin VB.TextBox txtUserName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2040
         TabIndex        =   7
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF8080&
         Caption         =   "Check1"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   255
      End
      Begin MSComCtl2.DTPicker DTPDateTo 
         Height          =   375
         Left            =   5040
         TabIndex        =   1
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   46333953
         CurrentDate     =   38368
      End
      Begin MSComCtl2.DTPicker DTPDateFrom 
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Format          =   46333953
         CurrentDate     =   38368
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF8080&
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4080
         TabIndex        =   3
         Top             =   360
         Width           =   615
      End
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      FormHeightDT    =   5130
      FormWidthDT     =   8460
      FormScaleHeightDT=   4620
      FormScaleWidthDT=   8340
   End
   Begin Crystal.CrystalReport crLog 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "USER LOGIN REPORT"
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
      Left            =   2040
      TabIndex        =   5
      Top             =   360
      Width           =   4230
   End
End
Attribute VB_Name = "frmLogReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdViewReport_Click()
'{Log_Details.LogID}

Dim strReport As String
strReport = App.Path & "\Reports\LogReport.rpt"

crLog.ReportFileName = App.Path & "\Reports\LogReport.rpt"
crLog.DiscardSavedData = True


If Check1.Value = 0 Then
   crLog.ReplaceSelectionFormula ("{Log_Details.Date}   >=#" & DTPDateFrom & "#  and {Log_Details.Date}  <=#" & DTPDateTo & "#")
ElseIf Check1.Value = 1 Then
    crLog.ReplaceSelectionFormula ("{Log_Details.Date}   >=#" & DTPDateFrom & "#  and {Log_Details.Date}  <=#" & DTPDateTo & "# and {Log_Details.UserName} = '" & txtUserName & "'")
End If

crLog.WindowState = crptMaximized
crLog.Action = 1


End Sub

