VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmInPatientMedReport 
   BackColor       =   &H00FF8080&
   Caption         =   "In Patient Medicine Report"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9210
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
   ScaleHeight     =   3375
   ScaleWidth      =   9210
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   8655
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   975
         Left            =   6960
         Picture         =   "frmInPatientMedReport.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cmbAdmitID 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton cmdView 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Report"
         Height          =   975
         Left            =   5400
         Picture         =   "frmInPatientMedReport.frx":0504
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "Patient Admission ID"
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
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   2295
      End
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   360
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      AutoCenterForm  =   -1  'True
      FormHeightDT    =   3885
      FormWidthDT     =   9330
      FormScaleHeightDT=   3375
      FormScaleWidthDT=   9210
   End
   Begin Crystal.CrystalReport crMedicine 
      Left            =   960
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "IN PATIENT MEDICINE REPORT"
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
      Left            =   1560
      TabIndex        =   4
      Top             =   480
      Width           =   6315
   End
End
Attribute VB_Name = "frmInPatientMedReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdView_Click()
If cmbAdmitID = "" Then
    MsgBox "Please select the Admisison ID", vbCritical
    Exit Sub
End If

Dim strReport As String
strReport = App.Path & "\Reports\InPatient\Medicine.rpt"


crMedicine.ReportFileName = App.Path & "\Reports\InPatient\Medicine.rpt"
crMedicine.DiscardSavedData = True
crMedicine.ReplaceSelectionFormula ("{InPatient_Orders.AdmissionID}  ='" & cmbAdmitID & "'")


crMedicine.WindowState = crptMaximized
crMedicine.Action = 1





End Sub

Private Sub Form_Load()
Dim rsAddID As Recordset
Set rsAddID = New ADODB.Recordset

rsAddID.Open "select * from Admission_Details", cnPatients, adOpenDynamic, adLockReadOnly

While rsAddID.EOF = False
cmbAdmitID.AddItem rsAddID(0)
rsAddID.MoveNext
Wend

rsAddID.Close


End Sub

