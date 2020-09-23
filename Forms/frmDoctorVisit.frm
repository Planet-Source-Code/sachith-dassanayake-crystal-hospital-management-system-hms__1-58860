VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmDoctorVisit 
   BackColor       =   &H00FF8080&
   Caption         =   "Doctor Visit Details"
   ClientHeight    =   10800
   ClientLeft      =   1590
   ClientTop       =   450
   ClientWidth     =   10890
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDoctorVisit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10800
   ScaleWidth      =   10890
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   9615
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   10215
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&OK"
         Height          =   855
         Left            =   3360
         Picture         =   "frmDoctorVisit.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   7800
         Width           =   1215
      End
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   855
         Left            =   5520
         Picture         =   "frmDoctorVisit.frx":5CB8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   7680
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6855
         Left            =   480
         TabIndex        =   4
         Top             =   360
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   12091
         _Version        =   393216
         AllowBigSelection=   0   'False
         HighLight       =   2
         SelectionMode   =   1
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF8080&
      Caption         =   "Visit Details"
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
      Height          =   5415
      Left            =   1080
      TabIndex        =   20
      Top             =   1200
      Width           =   6255
      Begin VB.ComboBox cmbDoctorID 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2160
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdAdmission 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         TabIndex        =   37
         Top             =   3300
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Visit_ID"
         Height          =   285
         Index           =   0
         Left            =   2520
         TabIndex        =   32
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Doctor_ID"
         Height          =   285
         Index           =   3
         Left            =   2520
         TabIndex        =   31
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Description"
         Height          =   285
         Index           =   6
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   3840
         Width           =   2415
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Prescription_ID"
         Height          =   285
         Index           =   7
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   4380
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton cmdPatient 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         TabIndex        =   26
         Top             =   2700
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdDoc 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5160
         TabIndex        =   25
         Top             =   2100
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cmbPatientID 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2760
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox cmbAdmitID 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3300
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtStatus 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         TabIndex        =   21
         Top             =   4860
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTPDate1 
         Height          =   315
         Left            =   2520
         TabIndex        =   33
         Top             =   900
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Format          =   45809665
         CurrentDate     =   38355
      End
      Begin MSComCtl2.DTPicker DTPTime1 
         Height          =   315
         Left            =   2520
         TabIndex        =   35
         Top             =   1500
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   556
         _Version        =   393216
         Format          =   45809666
         CurrentDate     =   38355
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Patient_ID"
         Height          =   285
         Index           =   4
         Left            =   2520
         TabIndex        =   30
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Admission_ID"
         Height          =   285
         Index           =   5
         Left            =   2520
         TabIndex        =   29
         Top             =   3300
         Width           =   2415
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Visit_Date"
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   34
         Top             =   900
         Width           =   2415
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Visit_Time"
         Height          =   285
         Index           =   2
         Left            =   2520
         TabIndex        =   36
         Top             =   1500
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Visit ID:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   46
         Top             =   420
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Visit Date:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   45
         Top             =   900
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Visit Time:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   44
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Doctor ID:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   43
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Patient ID:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   42
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Admission ID:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   41
         Top             =   3300
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Description:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   40
         Top             =   3900
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Prescription ID:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   600
         TabIndex        =   39
         Top             =   4395
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblStat 
         BackColor       =   &H00FF8080&
         Caption         =   "Status"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   38
         Top             =   4860
         Visible         =   0   'False
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   780
      Left            =   8640
      Picture         =   "frmDoctorVisit.frx":6104
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Update"
      Height          =   780
      Left            =   8640
      Picture         =   "frmDoctorVisit.frx":6608
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
      Height          =   780
      Left            =   8640
      Picture         =   "frmDoctorVisit.frx":6AB6
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Edit"
      Height          =   780
      Left            =   8640
      Picture         =   "frmDoctorVisit.frx":6F6F
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Add"
      Height          =   780
      Left            =   8640
      Picture         =   "frmDoctorVisit.frx":7414
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "Record Navigation"
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
      Height          =   1215
      Left            =   2040
      TabIndex        =   9
      Top             =   9240
      Width           =   7215
      Begin VB.CommandButton cmdFirst 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   360
         Picture         =   "frmDoctorVisit.frx":78C3
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   1065
         Picture         =   "frmDoctorVisit.frx":7D99
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   5520
         Picture         =   "frmDoctorVisit.frx":827A
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdLast 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   6240
         Picture         =   "frmDoctorVisit.frx":8755
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   1800
         TabIndex        =   14
         Top             =   480
         Width           =   3480
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Record Operations"
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
      Height          =   5415
      Left            =   8160
      TabIndex        =   5
      Top             =   1200
      Width           =   2055
      Begin VB.CommandButton cmdViewAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&View"
         Height          =   780
         Left            =   480
         Picture         =   "frmDoctorVisit.frx":8C2A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   780
         Left            =   480
         Picture         =   "frmDoctorVisit.frx":90E5
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   780
         Left            =   480
         Picture         =   "frmDoctorVisit.frx":95F1
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2760
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Height          =   2025
      Left            =   480
      TabIndex        =   0
      Top             =   6840
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   3572
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      TabAction       =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
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
      FormHeightDT    =   11310
      FormWidthDT     =   11010
      FormScaleHeightDT=   10800
      FormScaleWidthDT=   10890
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "VISIT DETAILS"
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
      Index           =   8
      Left            =   4080
      TabIndex        =   47
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "frmDoctorVisit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim IDType As Integer







Private Sub cmbAdmitID_Click()
Dim rsCheckStatus As Recordset
Dim rsAddData As Recordset

Set rsCheckStatus = New ADODB.Recordset
Set rsAddData = New ADODB.Recordset

rsCheckStatus.Open "Select * from In_Patient_Discharge where Admission_ID = '" & cmbAdmitID & "'", cnPatients, adOpenDynamic, adLockReadOnly
If rsCheckStatus.EOF = False Then
    txtStatus = "Discharged"
Else
    txtStatus = "Under Treatements"
End If

rsCheckStatus.Close






End Sub

Private Sub cmbPatientID_Click()
cmbAdmitID.clear
addAdmissionID cmbPatientID

If cmbAdmitID.ListCount = 0 Then
txtStatus = "Not Yet Admitted"
Else
    cmbAdmitID = cmbAdmitID.List(0)
End If

End Sub

Private Sub cmdAdmission_Click()
frmDisplayAdmissionDetails.Show
End Sub

Private Sub cmdExit_Click()


Frame1.Visible = False
Frame1.Caption = ""
If IDType = 0 Then
    cmbDoctorID.SetFocus
ElseIf IDType = 1 Then
    cmbPatientID.SetFocus
End If


End Sub

Private Sub cmdOK_Click()
If IDType = 0 Then
cmbDoctorID = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
cmbPatientID.SetFocus
End If
If IDType = 1 Then
cmbPatientID = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
txtFields(5).SetFocus
End If
Frame1.Caption = ""
Frame1.Visible = False

End Sub

Private Sub cmdDoc_Click()

Frame1.Caption = "Doctor Details"
Frame1.Visible = True
IDType = 0
Dim RowNo As Integer
Dim rsDoctor As New Recordset

Set rsDoctor = New ADODB.Recordset

rsDoctor.Open "Select * from Doctor_Details", cnPatients, adOpenDynamic, adLockReadOnly

If rsDoctor.EOF = False Then
RowNo = 1

With MSFlexGrid1
    .clear
    .Rows = 1
    .Cols = rsDoctor.Fields.Count
   
  

    While Not rsDoctor.EOF
        .Rows = .Rows + 1
        '.Row = .Rows - 1
        
    .TextMatrix(RowNo, 0) = rsDoctor(0)
    .TextMatrix(RowNo, 1) = rsDoctor(1)
    .TextMatrix(RowNo, 2) = rsDoctor(2)
    .TextMatrix(RowNo, 3) = rsDoctor(9)
    .TextMatrix(RowNo, 4) = rsDoctor(10)
    
    RowNo = RowNo + 1

      
    rsDoctor.MoveNext
    Wend
    
    
    .TextMatrix(0, 0) = "Doctor ID"
    .TextMatrix(0, 1) = "First Name"
    .TextMatrix(0, 2) = "Last Name"
    .TextMatrix(0, 3) = "Specialization"
    .TextMatrix(0, 4) = "Type"
    
    .FixedRows = 1
    .RowHeight(0) = .RowHeight(1) * 1.5
     Functions.SizeColumns MSFlexGrid1, Me
     .SetFocus
     .RowSel = 1
End With



rsDoctor.Close

End If





End Sub

Private Sub cmdPatient_Click()

Frame1.Caption = "Patient Details"
Frame1.Visible = True
IDType = 1
Dim RowNo As Integer
Dim rsPatient As New Recordset

Set rsPatient = New ADODB.Recordset
'create sql statement

'Set rsDocs.ActiveConnection = cnPatients

rsPatient.Open "Select * from In_Patient_Details", cnPatients, adOpenDynamic, adLockReadOnly

If rsPatient.EOF = False Then
RowNo = 1

With MSFlexGrid1
    .clear
    .Rows = 1
    .Cols = rsPatient.Fields.Count
  

    While Not rsPatient.EOF
        .Rows = .Rows + 1
        '.Row = .Rows - 1
    If Not rsPatient(0) = "" Then
        .TextMatrix(RowNo, 0) = rsPatient(0)
    End If
    If Not rsPatient(1) = "" Then
        .TextMatrix(RowNo, 1) = rsPatient(1)
    End If
    If Not rsPatient(2) = "" Then
        .TextMatrix(RowNo, 2) = rsPatient(2)
    End If
    If Not rsPatient(4) = "" Then
        .TextMatrix(RowNo, 3) = rsPatient(4)
    End If
    If Not rsPatient(5) = "" Then
        .TextMatrix(RowNo, 4) = rsPatient(5)
    End If
    
    RowNo = RowNo + 1

      
    rsPatient.MoveNext
    Wend
    
    
    .TextMatrix(0, 0) = "Patient ID"
    .TextMatrix(0, 1) = "First Name"
    .TextMatrix(0, 2) = "Last Name"
    .TextMatrix(0, 3) = "Sex"
    .TextMatrix(0, 4) = "NIC Number"
    
    .FixedRows = 1
    .RowHeight(0) = .RowHeight(1) * 1.5
     Functions.SizeColumns MSFlexGrid1, Me
End With



rsPatient.Close

End If



End Sub

Private Sub Form_Load()
  Call Functions.DisableMenu
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "SHAPE {select Visit_ID,Visit_Date,Visit_Time,Doctor_ID,Patient_ID,Admission_ID,Description,Prescription_ID from Visit_Details} AS ParentCMD APPEND ({select Visit_ID,Visit_Date,Visit_Time,Doctor_ID,Admission_ID,Patient_ID,Description,Prescription_ID from Visit_Details } AS ChildCMD RELATE Patient_ID TO Patient_ID) AS ChildCMD", cnPatients, adOpenDynamic, adLockPessimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next

  Set grdDataGrid.DataSource = adoPrimaryRS("ChildCMD").UnderlyingValue
    DTPDate1 = Date
    DTPTime1 = Time
    Call addPatientID
    Call addDoctorID
  
  mbDataChanged = False
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  'This will resize the grid when the form is resized
  'grdDataGrid.Width = Me.ScaleWidth
  'grdDataGrid.Height = Me.ScaleHeight - grdDataGrid.Top - 30 - picButtons.Height - picStatBox.Height
  'lblStatus.Width = Me.Width - 1500
  'cmdNext.Left = lblStatus.Width + 700
  'cmdLast.Left = cmdNext.Left + 340
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
  Call Functions.EnableMenu
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  'On Error GoTo AddErr
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    Call addVisitID
    lblStatus.Caption = "Add record"
    mbAddNewFlag = True
    SetButtons False
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
    
    If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "Confirm Delete") = vbNo Then
        Exit Sub
    End If
  
  With adoPrimaryRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  Set grdDataGrid.DataSource = Nothing
  adoPrimaryRS.Requery
  Set grdDataGrid.DataSource = adoPrimaryRS("ChildCMD").UnderlyingValue
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr

    DTPDate1 = txtFields(1)
    DTPTime1 = txtFields(2)
    cmbDoctorID = txtFields(3)
    cmbPatientID = txtFields(4)
    cmbAdmitID = txtFields(5)

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
    
    
    If txtStatus = "Discharged" Then
        MsgBox "Please Select the Correct Admission ID", vbCritical
        cmbAdmitID.SetFocus
        Exit Sub
    End If
        
    txtFields(1) = DTPDate1.Value
    txtFields(2) = DTPTime1.Value
    txtFields(3) = cmbDoctorID
    txtFields(4) = cmbPatientID
    txtFields(5) = cmbAdmitID

  If txtFields(3) = "" Or txtFields(4) = "" Or txtFields(5) = "" Then
    MsgBox "Please Enter all the relavant fields", vbCritical
    Exit Sub
  End If
  
    


  adoPrimaryRS.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'move to the new record
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  adoPrimaryRS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  adoPrimaryRS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    adoPrimaryRS.MoveLast
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveFirst
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
  
  txtFields(6).Locked = bVal
  DTPDate1.Visible = Not bVal
  DTPTime1.Visible = Not bVal
  cmbDoctorID.Visible = Not bVal
  cmbPatientID.Visible = Not bVal
  cmbAdmitID.Visible = Not bVal
  
  cmdDoc.Visible = Not bVal
  cmdPatient.Visible = Not bVal
  cmdAdmission.Visible = Not bVal
  
  txtStatus.Visible = Not bVal
  lblStat.Visible = Not bVal
  
  
End Sub

Private Sub addPatientID()

Dim rsAddPatientID As Recordset
Set rsAddPatientID = New ADODB.Recordset

rsAddPatientID.Open "select * from In_Patient_Details", cnPatients, adOpenDynamic, adLockReadOnly

While rsAddPatientID.EOF = False
cmbPatientID.AddItem rsAddPatientID(0)
rsAddPatientID.MoveNext
Wend


rsAddPatientID.Close


End Sub
Private Sub addDoctorID()

Dim rsAddDoctorID As Recordset
Set rsAddDoctorID = New ADODB.Recordset

rsAddDoctorID.Open "select * from Doctor_Details", cnPatients, adOpenDynamic, adLockReadOnly

While rsAddDoctorID.EOF = False
cmbDoctorID.AddItem rsAddDoctorID(0)
rsAddDoctorID.MoveNext
Wend


rsAddDoctorID.Close

End Sub

Private Sub addVisitID()
Dim VisitID As String
Dim rsAddVisitID As Recordset

    VisitID = Functions.UID(6, "DocVis_")
    
    Set rsAddVisitID = New ADODB.Recordset
    rsAddVisitID.Open "Select * from Visit_Details", cnPatients, adOpenKeyset, adLockPessimistic
              
            While rsAddVisitID.EOF = False
                If rsAddVisitID(0) = VisitID Then
                    VisitID = Functions.UID(6, "DocVis_")
                    rsAddVisitID.MoveFirst
                End If
                rsAddVisitID.MoveNext
            Wend
             
txtFields(0) = VisitID
End Sub
Private Sub MSFlexGrid1_DblClick()


    If IDType = 0 Then
        cmbDoctorID = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
    End If
    If IDType = 1 Then
        cmbPatientID = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
    End If
    cmbPatientID.SetFocus
    Frame1.Visible = False
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If IDType = 0 Then
        cmbDoctorID = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
        cmbPatientID.SetFocus
    End If
    If IDType = 1 Then
        cmbPatientID = MSFlexGrid1.TextMatrix(MSFlexGrid1.Row, 0)
        txtFields(5).SetFocus
    End If
    Frame1.Visible = False
End If

If KeyAscii = 27 Then
    If IDType = 0 Then
        cmbDoctorID.SetFocus
    ElseIf IDType = 1 Then
        cmbPatientID.SetFocus
    End If

    Frame1.Visible = False
End If

End Sub

Private Sub addAdmissionID(AddID As String)
Dim rsAddID As Recordset
Set rsAddID = New ADODB.Recordset

rsAddID.Open "Select * from Admission_Details where Patient_ID = '" & AddID & "'", cnPatients, adOpenDynamic, adLockReadOnly

While rsAddID.EOF = False
cmbAdmitID.AddItem rsAddID(0)
rsAddID.MoveNext
Wend

rsAddID.Close
End Sub
