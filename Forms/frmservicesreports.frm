VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmservicesreports 
   BackColor       =   &H00FF8080&
   Caption         =   "Service Reports"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4965
   ScaleWidth      =   11835
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9840
      Picture         =   "frmservicesreports.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "Services Reports"
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
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   11295
      Begin VB.CommandButton cmd_ser_app 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Services &Appointment Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   7440
         Picture         =   "frmservicesreports.frx":0504
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmd_ser_shedual 
         BackColor       =   &H00FFFFFF&
         Caption         =   "S&ervices Schedule Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   5040
         Picture         =   "frmservicesreports.frx":0B2A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmd_ser_total 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total &Services Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   360
         Picture         =   "frmservicesreports.frx":10A2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmd_ser_bill_payment 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Services &Bill Payment Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   2640
         Picture         =   "frmservicesreports.frx":161A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmd_ser_bill 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Services &Bill Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   9600
         Picture         =   "frmservicesreports.frx":1BEC
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "SERVICE REPORTS"
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
      TabIndex        =   7
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmservicesreports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_ser_app_Click()
Report.ReportFileName = App.Path & "\Reports\services\ServicesAppointments.rpt"
Report.DiscardSavedData = True
Report.WindowState = crptMaximized
Report.Action = 1
End Sub

Private Sub cmd_ser_bill_Click()
Report.ReportFileName = App.Path & "\Reports\services\ServicesAppointmentBill.rpt"
Report.DiscardSavedData = True
Report.WindowState = crptMaximized
Report.Action = 1
End Sub

Private Sub cmd_ser_bill_payment_Click()
Report.ReportFileName = App.Path & "\Reports\services\ServicesAppointmentBillpayment.rpt "
Report.DiscardSavedData = True
Report.WindowState = crptMaximized
Report.Action = 1
End Sub

Private Sub cmd_ser_shedual_Click()
Report.ReportFileName = App.Path & "\Reports\services\servicesschedule.rpt"
Report.DiscardSavedData = True
Report.WindowState = crptMaximized
Report.Action = 1
End Sub

Private Sub cmd_ser_total_Click()
Report.ReportFileName = App.Path & "\Reports\services\services.rpt"
Report.DiscardSavedData = True
Report.ReplaceSelectionFormula ("")
Report.WindowState = crptMaximized
Report.Action = 1
End Sub
