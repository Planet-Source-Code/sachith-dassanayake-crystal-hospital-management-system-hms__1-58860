VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmemployeereports 
   BackColor       =   &H00FF8080&
   Caption         =   "Reports"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12210
   Icon            =   "frmemployeereports.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   12210
   Begin VB.Frame frm_emp_reports 
      BackColor       =   &H00FF8080&
      Caption         =   "Employee Reports"
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
      Height          =   4215
      Left            =   480
      TabIndex        =   0
      Top             =   1080
      Width           =   11295
      Begin VB.CommandButton cmd_emp_total 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Display &Total Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   360
         Picture         =   "frmemployeereports.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmd_emp_only 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Employee Only"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1680
         Picture         =   "frmemployeereports.frx":5CD1
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmd_emp_doconly 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Doctor Only"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1680
         Picture         =   "frmemployeereports.frx":615A
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FF8080&
         Caption         =   "Custom Reports"
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
         Height          =   3495
         Left            =   3120
         TabIndex        =   2
         Top             =   360
         Width           =   7815
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FF8080&
            Caption         =   "Employee"
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
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00FF8080&
            Caption         =   "Doctor"
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
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox txtfrom 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   9
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox txtto 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   8
            Top             =   1440
            Width           =   2175
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00FF8080&
            ForeColor       =   &H00FFFFFF&
            Height          =   1455
            Left            =   2400
            TabIndex        =   4
            Top             =   480
            Width           =   2655
            Begin VB.OptionButton Option2 
               BackColor       =   &H00FF8080&
               Caption         =   "By Net Salary"
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
               Index           =   3
               Left            =   120
               TabIndex        =   7
               Top             =   600
               Width           =   1815
            End
            Begin VB.OptionButton Option2 
               BackColor       =   &H00FF8080&
               Caption         =   "By Gross Salary"
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
               Index           =   2
               Left            =   120
               TabIndex        =   6
               Top             =   960
               Width           =   1935
            End
            Begin VB.OptionButton Option2 
               BackColor       =   &H00FF8080&
               Caption         =   "By Basic Salary"
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
               Index           =   0
               Left            =   120
               TabIndex        =   5
               Top             =   240
               Width           =   1815
            End
         End
         Begin VB.CommandButton cmd_custom 
            BackColor       =   &H00FFFFFF&
            Caption         =   "&Generate"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   6000
            Picture         =   "frmemployeereports.frx":65E3
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
            Caption         =   "Between"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   5880
            TabIndex        =   13
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
            Caption         =   "And"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   6000
            TabIndex        =   12
            Top             =   1080
            Width           =   735
         End
      End
      Begin VB.CommandButton cmd_per_doc_only 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Permanant Doctor Only"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   360
         Picture         =   "frmemployeereports.frx":671B
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1800
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Report 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      FormHeightDT    =   6210
      FormWidthDT     =   12330
      FormScaleHeightDT=   5700
      FormScaleWidthDT=   12210
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "EMPLOYEE REPORTS"
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
      Left            =   4200
      TabIndex        =   17
      Top             =   240
      Width           =   4125
   End
End
Attribute VB_Name = "frmemployeereports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public flag1 As String
Public flag2 As String

Private Sub cmd_emp_doconly_Click()
Report.ReportFileName = App.Path & "\Reports\Employee\total_doctor.rpt"
Report.DiscardSavedData = True
Report.WindowState = crptMaximized
Report.Action = 1
End Sub

Private Sub cmd_emp_only_Click()
Report.ReportFileName = App.Path & "\Reports\Employee\employee.rpt"
Report.DiscardSavedData = True
Report.ReplaceSelectionFormula ("")
Report.WindowState = crptMaximized
Report.Action = 1
End Sub

Private Sub cmd_emp_total_Click()
Report.ReportFileName = App.Path & "\Reports\Employee\employee.rpt"
Report.DiscardSavedData = True
Report.ReplaceSelectionFormula ("")
Report.WindowState = crptMaximized
Report.Action = 1
End Sub

Private Sub cmd_per_doc_only_Click()
Report.ReportFileName = App.Path & "\Reports\Employee\permenant Doctor.rpt"
Report.DiscardSavedData = True
Report.WindowState = crptMaximized
Report.Action = 1
End Sub



Private Sub Option1_Click(Index As Integer)

If Option1(0).Value = True Then
flag1 = "employee"
ElseIf Option1(1).Value = True Then
flag1 = "doctor"
End If

End Sub

Private Sub Option2_Click(Index As Integer)
If Option2(0).Value = True Then
flag2 = "bsal"
ElseIf Option2(1).Value = True Then
flag2 = "net"
ElseIf Option2(2).Value = True Then
flag2 = "gross"
End If
End Sub
Private Sub cmd_custom_Click()

'{SLIP.gross} >= 0  and {SLIP.gross} <= 1000000

'{SLIP.net} >= 0  and  {SLIP.net} <= 100000
If flag1 = "employee" And flag2 = "" Then

Report.ReportFileName = App.Path & "\Reports\Employee\Emplyee_salary.rpt"
Report.DiscardSavedData = True
Report.WindowState = crptMaximized
Report.Action = 1

ElseIf flag1 = "doctor" And flag2 = "" Then

Report.ReportFileName = App.Path & "\Reports\Employee\perdoctorsal.rpt"
Report.DiscardSavedData = True
Report.WindowState = crptMaximized
Report.Action = 1

ElseIf flag1 = "employee" And flag2 = "bsal" Then

Report.ReportFileName = App.Path & "\Reports\Employee\Emplyee_salary.rpt"
Report.DiscardSavedData = True
Report.ReplaceSelectionFormula ("{Employee.bsal}  >= " & Val(txtfrom) & " and  {Employee.bsal}<= " & Val(txtto) & "")
Report.WindowState = crptMaximized
Report.Action = 1

ElseIf flag1 = "employee" And flag2 = "net" Then

Report.ReportFileName = App.Path & "\Reports\Employee\Emplyee_salary.rpt"
Report.DiscardSavedData = True
Report.ReplaceSelectionFormula ("{SLIP.net} >= " & Val(txtfrom) & "  and  {SLIP.net} <= " & Val(txtto) & "")
Report.WindowState = crptMaximized
Report.Action = 1

ElseIf flag1 = "employee" And flag2 = "gross" Then

Report.ReportFileName = App.Path & "\Reports\Employee\Emplyee_salary.rpt"
Report.DiscardSavedData = True
Report.ReplaceSelectionFormula ("{SLIP.gross} >= " & Val(txtfrom) & "  and {SLIP.gross} <= " & Val(txtto) & "")
Report.WindowState = crptMaximized
Report.Action = 1

ElseIf flag1 = "doctor" And flag2 = "bsal" Then

Report.ReportFileName = App.Path & "\Reports\Employee\perdoctorsal.rpt"
Report.DiscardSavedData = True
Report.ReplaceSelectionFormula ("{Doctor_salary.basic} >= " & Val(txtfrom) & " and {Doctor_salary.basic} <= " & Val(txtto) & "")
Report.WindowState = crptMaximized
Report.Action = 1

ElseIf flag1 = "doctor" And flag2 = "net" Then

Report.ReportFileName = App.Path & "\Reports\Employee\perdoctorsal.rpt"
Report.DiscardSavedData = True
Report.ReplaceSelectionFormula ("{Doctor_salary.net} >= " & Val(txtfrom) & " and {Doctor_salary.net} <= " & Val(txtto) & "")
Report.WindowState = crptMaximized
Report.Action = 1


ElseIf flag1 = "doctor" And flag2 = "gross" Then

Report.ReportFileName = App.Path & "\Reports\Employee\perdoctorsal.rpt"
Report.DiscardSavedData = True
Report.ReplaceSelectionFormula ("{Doctor_salary.gross} >= " & Val(txtfrom) & " and {Doctor_salary.gross} <= " & Val(txtto) & "")
Report.WindowState = crptMaximized
Report.Action = 1

End If


End Sub
