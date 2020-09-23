VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmAdmissionDetails 
   BackColor       =   &H00FF8080&
   Caption         =   "Admission Details"
   ClientHeight    =   10560
   ClientLeft      =   2985
   ClientTop       =   1845
   ClientWidth     =   11475
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAdmissionDetails.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10560
   ScaleWidth      =   11475
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   780
      Left            =   4920
      Picture         =   "frmAdmissionDetails.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   8280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Update"
      Height          =   780
      Left            =   3720
      Picture         =   "frmAdmissionDetails.frx":5CE6
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   8280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&View"
      Height          =   780
      Left            =   4920
      Picture         =   "frmAdmissionDetails.frx":6194
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   9240
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   780
      Left            =   6120
      Picture         =   "frmAdmissionDetails.frx":664F
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   9240
      Width           =   1095
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      Height          =   780
      Left            =   7320
      Picture         =   "frmAdmissionDetails.frx":6B5B
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
      Height          =   780
      Left            =   6120
      Picture         =   "frmAdmissionDetails.frx":7001
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Edit"
      Height          =   780
      Left            =   4920
      Picture         =   "frmAdmissionDetails.frx":74BA
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Add"
      Height          =   780
      Left            =   3720
      Picture         =   "frmAdmissionDetails.frx":795F
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   8280
      Width           =   1095
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   2640
      Picture         =   "frmAdmissionDetails.frx":7E0E
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   7080
      UseMaskColor    =   -1  'True
      Width           =   585
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   3345
      Picture         =   "frmAdmissionDetails.frx":82E4
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   7080
      UseMaskColor    =   -1  'True
      Width           =   585
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   7800
      Picture         =   "frmAdmissionDetails.frx":87C5
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   7080
      UseMaskColor    =   -1  'True
      Width           =   585
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   8520
      Picture         =   "frmAdmissionDetails.frx":8CA0
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   7080
      UseMaskColor    =   -1  'True
      Width           =   585
   End
   Begin VB.TextBox txtBedAvail 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   5280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdCheckBed 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bed Details"
      Height          =   855
      Left            =   9240
      Picture         =   "frmAdmissionDetails.frx":9175
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "Room / Ward"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   240
      TabIndex        =   28
      Top             =   4560
      Visible         =   0   'False
      Width           =   5295
      Begin VB.CommandButton cmdViewWard 
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
         Height          =   255
         Left            =   4680
         TabIndex        =   34
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdViewRoom 
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
         Height          =   255
         Left            =   4680
         TabIndex        =   33
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cmbRoomID 
         DataSource      =   "Room_ID"
         Height          =   315
         ItemData        =   "frmAdmissionDetails.frx":966B
         Left            =   2040
         List            =   "frmAdmissionDetails.frx":966D
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   240
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "Room"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "Ward"
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cmbWardID 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   720
         Visible         =   0   'False
         Width           =   2535
      End
   End
   Begin VB.ComboBox cmbBedID 
      Height          =   315
      Left            =   7680
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   4800
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ComboBox cmbDoctorID 
      Height          =   315
      Left            =   7680
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   3240
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ComboBox cmbPatientID 
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   2040
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdInPatientID 
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
      Height          =   255
      Left            =   5040
      TabIndex        =   24
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdInGuardianID 
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
      Height          =   255
      Left            =   10440
      TabIndex        =   23
      Top             =   2040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cmbGuardianID 
      Height          =   315
      Left            =   7680
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   2040
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdRefDoc 
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
      Height          =   255
      Left            =   10440
      TabIndex        =   19
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdBedDetails 
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
      Height          =   255
      Left            =   10560
      TabIndex        =   18
      Top             =   4800
      Width           =   495
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Bed_ID"
      Height          =   285
      Index           =   8
      Left            =   7680
      TabIndex        =   17
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Room_Ward_ID"
      Height          =   285
      Index           =   7
      Left            =   2280
      TabIndex        =   15
      Top             =   4920
      Width           =   2655
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ref_Doctor"
      Height          =   285
      Index           =   6
      Left            =   7680
      TabIndex        =   13
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Emergency"
      Height          =   285
      Index           =   5
      Left            =   2280
      TabIndex        =   11
      Top             =   3225
      Width           =   2655
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Admission_Time"
      Height          =   285
      Index           =   4
      Left            =   7680
      TabIndex        =   9
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Guardian_ID"
      Height          =   285
      Index           =   2
      Left            =   7680
      TabIndex        =   5
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Patient_ID"
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   2055
      Width           =   2655
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Admission_ID"
      Height          =   285
      Index           =   0
      Left            =   2280
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker DTPTime 
      Height          =   285
      Left            =   7680
      TabIndex        =   20
      Top             =   2640
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20709378
      CurrentDate     =   38327
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   285
      Left            =   2280
      TabIndex        =   21
      Top             =   2640
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   20709377
      CurrentDate     =   38327
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Admission_Date"
      Height          =   285
      Index           =   3
      Left            =   2280
      TabIndex        =   7
      Top             =   2700
      Width           =   2655
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      FormHeightDT    =   11070
      FormWidthDT     =   11595
      FormScaleHeightDT=   10560
      FormScaleWidthDT=   11475
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "ADMISSION DETAILS"
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
      Index           =   9
      Left            =   3960
      TabIndex        =   51
      Top             =   360
      Width           =   4290
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      Height          =   2415
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Width           =   11175
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   4170
      TabIndex        =   48
      Top             =   7200
      Width           =   3360
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   2175
      Left            =   2520
      Shape           =   4  'Rounded Rectangle
      Top             =   8040
      Width           =   6855
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   975
      Left            =   2160
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Width           =   7455
   End
   Begin VB.Label lblBedStat 
      BackColor       =   &H00FF8080&
      Caption         =   "Bed Availability"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   36
      Top             =   5280
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   2415
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   11175
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Bed ID:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   5880
      TabIndex        =   16
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Room/Ward ID:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   14
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Reffered Doctor:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   5760
      TabIndex        =   12
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Emergency:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   10
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Admission Time:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   5760
      TabIndex        =   8
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Admission Date:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Guardian ID:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   4
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Patient ID:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Admission ID:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   1815
   End
End
Attribute VB_Name = "frmAdmissionDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim BMngID As String
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean


Private Function ChkBedAvailability(bed As String) As Boolean


Dim rsBedAvail As Recordset
Set rsBedAvail = New ADODB.Recordset

rsBedAvail.Open "select * from Bed_Details where Bed_ID = '" & cmbBedID.Text & "'", cnPatients, adOpenDynamic, adLockPessimistic

If rsBedAvail.EOF = True Then
    Debug.Print "The Selected Bed is Available"
    ChkBedAvailability = True
ElseIf rsBedAvail![available] = False Then
    ChkBedAvailability = False
ElseIf rsBedAvail![available] = True Then
    ChkBedAvailability = True
End If

rsBedAvail.Close




End Function



Private Sub cmbBedID_Click()

Dim result As Boolean
result = ChkBedAvailability(cmbBedID.Text)
If result = True Then
    txtBedAvail = "Available"
ElseIf result = False Then
    txtBedAvail = "Not Available"
End If

Debug.Print "Bed Availability = " & result
End Sub

Private Sub cmbRoomID_Click()

Dim rsBedID As Recordset
Dim rsRoomID As Recordset

Set rsRoomID = New ADODB.Recordset
Set rsBedID = New ADODB.Recordset

rsRoomID.Open " select Room_ID from Room_Details where Room_Type = '" & cmbRoomID.Text & "'", cnPatients, adOpenDynamic, adLockPessimistic
cmbBedID.clear
While rsRoomID.EOF = False
    rsBedID.Open "Select Bed_ID from Bed_Details where Room_Ward_ID= '" & rsRoomID(0) & "'", cnPatients, adOpenDynamic, adLockPessimistic
    
    Debug.Print rsBedID.RecordCount
    
    While rsBedID.EOF = False
        cmbBedID.AddItem (rsBedID(0))
        cmbBedID.Text = rsBedID(0)
        rsBedID.MoveNext
     Wend

    
    If rsBedID.EOF = False Then
        cmbBedID.AddItem (rsBedID(0))
        cmbBedID.Text = rsBedID(0)
    End If
    
    
    rsBedID.Close
    
    rsRoomID.MoveNext
Wend

rsRoomID.Close
cmbBedID_Click
cmbBedID.Enabled = True

End Sub




Private Sub cmbWardID_Click()


Dim rsBedID As Recordset
Dim rsWardID As Recordset

Set rsWardID = New ADODB.Recordset
Set rsBedID = New ADODB.Recordset

rsWardID.Open " select Ward_ID from Ward_Details where Ward_Name = '" & cmbWardID.Text & "'", cnPatients, adOpenDynamic, adLockPessimistic
cmbBedID.clear
While rsWardID.EOF = False
    rsBedID.Open "Select Bed_ID from Bed_Details where Room_Ward_ID= '" & rsWardID(0) & "'", cnPatients, adOpenDynamic, adLockPessimistic
    Debug.Print rsBedID.RecordCount
    
   If rsBedID.EOF = True Then
    txtBedAvail = "Not Available"
    Exit Sub
   End If
    
     While rsBedID.EOF = False
        cmbBedID.AddItem (rsBedID(0))
        cmbBedID.Text = rsBedID(0)
        rsBedID.MoveNext
     Wend
    
    
    
    rsBedID.Close
    
    rsWardID.MoveNext
Wend

rsWardID.Close
cmbBedID_Click
cmbBedID.Enabled = True

End Sub

Private Sub cmdBedDetails_Click()
frmBedDetails.Show
End Sub

Private Sub cmdCheckBed_Click()
frmBEDDisplay.Show
End Sub



Private Sub cmdInGuardianID_Click()
frmDisplayGuardian.Show
End Sub

Private Sub cmdInPatientID_Click()
frmDisplayInPatient.Show
End Sub

Private Sub cmdRefDoc_Click()
frmDoctorDetails.Show
End Sub

Private Sub cmdView_Click()
frmDisplayAdmissionDetails.Show

End Sub

Private Sub cmdViewRoom_Click()
frmRoomDetails.Show
End Sub

Private Sub cmdViewWard_Click()
frmWardDetails.Show
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
Call Functions.DisableMenu
End Sub



Private Sub Form_Load()

     Me.WindowState = vbMaximized


   Call Functions.DisableMenu
   Set adoPrimaryRS = New Recordset
   adoPrimaryRS.Open "select Admission_ID,Patient_ID,Guardian_ID,Admission_Date,Admission_Time,Emergency,Ref_Doctor,Room_Ward_ID,Bed_ID from Admission_Details", cnPatients, adOpenDynamic, adLockOptimistic

    DTPDate = Date
    DTPTime = Time
 
  
    
    Dim rsPID As Recordset
    Set rsPID = New ADODB.Recordset
    rsPID.Open "select * from In_Patient_Details", cnPatients, adOpenStatic, adLockPessimistic
    While rsPID.EOF = False
        cmbPatientID.AddItem rsPID(0)
        rsPID.MoveNext
    Wend
   
    rsPID.MoveLast
    cmbPatientID.Text = rsPID(0)
    rsPID.Close
    
    Dim rsGID As Recordset
    Set rsGID = New ADODB.Recordset
    rsGID.Open "select * from Guardian_Details", cnPatients, adOpenStatic, adLockPessimistic
    While rsGID.EOF = False
        cmbGuardianID.AddItem rsGID(0)
        rsGID.MoveNext
    Wend
   
    rsGID.MoveLast
    cmbGuardianID.Text = rsGID(0)
    rsGID.Close
    
    Dim rsDID As Recordset
    Set rsDID = New ADODB.Recordset
    rsDID.Open "select * from Doctor_Details", cnPatients, adOpenStatic, adLockPessimistic
    While rsDID.EOF = False
        cmbDoctorID.AddItem rsDID(0)
        rsDID.MoveNext
    Wend
   
    rsDID.MoveLast
    cmbDoctorID.Text = rsDID(0)
    rsDID.Close
    
    
    Option1_Click (1)
    Option1_Click (0)
    cmbRoomID_Click
    

    

  
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
    oText.Locked = True
    oText.Appearance = 0
  Next

  mbDataChanged = False
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

Public Sub cmdAdd_Click()
  On Error GoTo AddErr
  
Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    oText.Locked = False
Next
  txtFields(0).Locked = True

  
  
  
    Dim rsAdmit As Recordset
    Dim AdmID As String
    
    
    AdmID = Functions.UID(6, "AdmID_")
    Set rsAdmit = New ADODB.Recordset
    rsAdmit.Open "Select * from Admission_Details", cnPatients, adOpenKeyset, adLockPessimistic
              
            While rsAdmit.EOF = False
                If rsAdmit(0) = RoomID Then
                    AdmID = Functions.UID(6, "AdmID_")
                    rsAdmit.MoveFirst
                End If
                rsAdmit.MoveNext
            Wend
            rsAdmit.Close
            

   
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    txtFields(0) = AdmID
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
  
  Dim rsDelBed As Recordset
  Set rsDelBed = New ADODB.Recordset
  
  If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "Delete Admission Record") = vbYes Then
  
    rsDelBed.Open "select * from Bed_Details where Bed_ID = '" & txtFields(8) & "'", cnPatients, adOpenDynamic, adLockOptimistic
  
  If rsDelBed.EOF = False Then
    
    rsDelBed(2) = True ' Bed availability true
    rsDelBed(3) = ""
    rsDelBed.Update
    rsDelBed.Close
  
  With adoPrimaryRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  
  Exit Sub
  End If
  
  
  Else
    Exit Sub
  End If
  
  
  
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  adoPrimaryRS.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    
    oText.Locked = False
  Next
   txtFields(0).Locked = True
  
  
  
  
  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next
  
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    
    oText.Locked = True
  Next
   txtFields(0).Locked = True

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  rsAddBedMng.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
  Dim oText As TextBox
  Dim rsAdmit As Recordset
  Set rsAdmit = New ADODB.Recordset
  Dim rsDis As Recordset
  Set rsDis = New ADODB.Recordset
  
  rsAdmit.Open "Select * from Admission_Details where Patient_ID= '" & cmbPatientID & "'", cnPatients, adOpenDynamic, adLockReadOnly

        rsDis.Open "Select * from In_Patient_Discharge where Patient_ID = '" & cmbPatientID & "'", cnPatients, adOpenDynamic, adLockReadOnly
        If rsAdmit.RecordCount > rsDis.RecordCount Then
            MsgBox "Patient has already admitted to the hospital", vbCritical
            Exit Sub
        End If
        If rsDis.RecordCount > rsAdmit.RecordCount Then
            MsgBox "Database Error.. Patient has discharged with out addmitting " & vbCrLf & "Please contact the database administrator", vbCritical
            Exit Sub
        End If
    
  rsAdmit.Close
 
  Dim rsRoom_Ward_ID As Recordset
  Dim strRoom_Ward_ID
  Set rsRoom_Ward_ID = New ADODB.Recordset
  
  rsRoom_Ward_ID.Open "select Room_Ward_ID from Bed_Details where Bed_ID= '" & cmbBedID & "' ", cnPatients, adOpenDynamic, adLockPessimistic
  
  If rsRoom_Ward_ID.EOF = True Then
    MsgBox "Room or Ward ID Not Found"
    Exit Sub
  Else
  strRoom_Ward_ID = rsRoom_Ward_ID(0)
  End If
  
  rsRoom_Ward_ID.Close
  
  
  txtFields(1) = cmbPatientID.Text
  txtFields(2) = cmbGuardianID.Text
  txtFields(3) = DTPDate.Value
  txtFields(4) = Format(DTPTime.Value, "short time")
  txtFields(6) = cmbDoctorID
  txtFields(7) = strRoom_Ward_ID
  txtFields(8) = cmbBedID

  ' Add Data to admission table and bed_details table
  
  Dim result As Boolean
  
  result = ChkBedAvailability(cmbBedID)
  If result = True Then
    'MsgBox "Bed is Available... , Continue the Admission Process"
  ElseIf result = False Then
    MsgBox "Bed is Not Available... , Please Select a different Bed"
   Exit Sub
  End If
  
  Dim rsAddBed As Recordset
  Set rsAddBed = New ADODB.Recordset
  
  rsAddBed.Open "select * from Bed_Details where Bed_ID = '" & cmbBedID & "'", cnPatients, adOpenDynamic, adLockOptimistic
  
  If rsAddBed.EOF = False Then
    adoPrimaryRS.UpdateBatch adAffectAll
    rsAddBed(2) = False
    rsAddBed(3) = txtFields(0)
    rsAddBed.Update
    rsAddBed.Close
  
  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'move to the new record
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  
  
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    oText.Locked = False
  Next
  txtFields(0).Locked = True

  Exit Sub
  
  End If
  
  
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    oText.Locked = False
  Next
    txtFields(0).Locked = True
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
  
  cmbPatientID.Visible = Not bVal
  cmbGuardianID.Visible = Not bVal
  cmbDoctorID.Visible = Not bVal
  cmbRoomID.Visible = Not bVal
  cmbWardID.Visible = Not bVal
  cmbBedID.Visible = Not bVal
  
  cmdInGuardianID.Visible = Not bVal
  cmdInPatientID.Visible = Not bVal
  cmdRefDoc.Visible = Not bVal
  Frame3.Visible = Not bVal
  Option1(0).Visible = Not bVal
  Option1(1).Visible = Not bVal
  
  DTPDate.Visible = Not bVal
  DTPTime.Visible = Not bVal
  
  txtFields(1).Visible = bVal
  txtFields(2).Visible = bVal
  txtFields(3).Visible = bVal
  txtFields(4).Visible = bVal
  txtFields(6).Visible = bVal
  txtFields(7).Visible = bVal
  
  txtBedAvail.Visible = Not bVal
  lblBedStat.Visible = Not bVal
  
  cmdView.Visible = bVal
  
  
End Sub

Private Sub Option1_Click(Index As Integer)
Dim rsRoom As Recordset
Dim rsWard As Recordset

Set rsRoom = New ADODB.Recordset
Set rsWard = New ADODB.Recordset
cmbRoomID.clear
cmbWardID.clear


Select Case (Index)
        
        Case "0" ' Room ID's
            cmbRoomID.Enabled = True
            cmbWardID.Enabled = False
            rsRoom.Open "select distinct(Room_type) from Room_Types", cnPatients, adOpenDynamic, adLockPessimistic
            While rsRoom.EOF = False
                cmbRoomID.AddItem (rsRoom(0))
                rsRoom.MoveNext
            Wend
            rsRoom.MoveFirst
            cmbRoomID.Text = rsRoom(0)
            rsRoom.Close
            cmbRoomID_Click
        
        Case "1" ' Ward ID
            
            cmbRoomID.Enabled = False
            cmbWardID.Enabled = True
          
            rsWard.Open "select distinct(Ward_Name) from Ward_Details", cnPatients, adOpenDynamic, adLockPessimistic
            While rsWard.EOF = False
                cmbWardID.AddItem (rsWard(0))
                rsWard.MoveNext
            Wend
            rsWard.MoveFirst
            cmbWardID.Text = rsWard(0)
            rsWard.Close
            cmbWardID_Click
            

        Case Else 'None
            
End Select

End Sub

