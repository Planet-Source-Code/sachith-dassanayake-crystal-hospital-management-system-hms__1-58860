VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmInPatientServiceReport 
   BackColor       =   &H00FF8080&
   Caption         =   "View Medical Services Report"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8700
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
   ScaleHeight     =   2085
   ScaleWidth      =   8700
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   120
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      AutoCenterForm  =   -1  'True
      FormHeightDT    =   2595
      FormWidthDT     =   8820
      FormScaleHeightDT=   2085
      FormScaleWidthDT=   8700
   End
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Report"
      Height          =   975
      Left            =   5400
      Picture         =   "frmInPatientServiceReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   975
      Left            =   6840
      Picture         =   "frmInPatientServiceReport.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin Crystal.CrystalReport crService 
      Left            =   360
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.ComboBox cmbAdmitID 
      Height          =   315
      ItemData        =   "frmInPatientServiceReport.frx":0A90
      Left            =   2640
      List            =   "frmInPatientServiceReport.frx":0A92
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   240
      Top             =   360
      Width           =   8055
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
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "frmInPatientServiceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Click()

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdView_Click()
If cmbAdmitID = "" Then
    MsgBox "Please select the Admisison ID", vbCritical
    Exit Sub
End If

Dim strReport As String
strReport = App.Path & "\Reports\InPatient\Services.rpt"


crService.ReportFileName = App.Path & "\Reports\InPatient\Services.rpt"
crService.DiscardSavedData = True
crService.ReplaceSelectionFormula ("{InPatient_Services.AdmissionID}  ='" & cmbAdmitID & "'")


crService.WindowState = crptMaximized
crService.Action = 1
End Sub

Private Sub Form_Load()
Dim rsadd As Recordset
Set rsadd = New ADODB.Recordset

rsadd.Open "Select * from Admission_Details", cnPatients, adOpenDynamic, adLockReadOnly

While rsadd.EOF = False
cmbAdmitID.AddItem rsadd(0)
rsadd.MoveNext
Wend

rsadd.Close

End Sub
