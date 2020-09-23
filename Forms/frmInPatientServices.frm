VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmInPatientServices 
   BackColor       =   &H00FF8080&
   Caption         =   "In Patient Lab Exams and Medical Services"
   ClientHeight    =   8790
   ClientLeft      =   1020
   ClientTop       =   1515
   ClientWidth     =   12345
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInPatientServices.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   12345
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Patient Details"
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
      Height          =   5055
      Left            =   480
      TabIndex        =   14
      Top             =   2160
      Width           =   5775
      Begin VB.CommandButton cmdViewAdmission 
         Caption         =   "..."
         Height          =   255
         Left            =   5040
         TabIndex        =   33
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton cmdViewPatient 
         Caption         =   "..."
         Height          =   255
         Left            =   5040
         TabIndex        =   32
         Top             =   360
         Width           =   495
      End
      Begin VB.ComboBox cmbAdmitID 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Height          =   1260
         Left            =   1800
         TabIndex        =   22
         Top             =   3240
         Width           =   3105
      End
      Begin VB.TextBox txtPatientName 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1800
         TabIndex        =   16
         Top             =   960
         Width           =   3105
      End
      Begin VB.ComboBox cmbInPatientID 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   360
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker DTPSDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   34
         Top             =   2040
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20709377
         CurrentDate     =   38354
      End
      Begin MSComCtl2.DTPicker DTPSTime 
         Height          =   375
         Left            =   1800
         TabIndex        =   35
         Top             =   2640
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20709378
         CurrentDate     =   38354
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discription"
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
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   3255
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name of Patient"
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
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   975
         Width           =   1560
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Treatment Time"
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
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   2775
         Width           =   1560
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Admission ID"
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
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1575
         Width           =   1290
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Treatment Date"
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
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   2175
         Width           =   1530
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Code"
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
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1230
      End
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CLOSE"
      Height          =   855
      Left            =   6840
      Picture         =   "frmInPatientServices.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7680
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      Height          =   855
      Left            =   5160
      Picture         =   "frmInPatientServices.frx":5CE6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox txtBillID 
      Height          =   285
      Left            =   5640
      TabIndex        =   8
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Service Details"
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
      Height          =   5055
      Left            =   6840
      TabIndex        =   1
      Top             =   2160
      Width           =   5055
      Begin VB.CommandButton cmdViewService 
         Caption         =   "..."
         Height          =   255
         Left            =   4080
         TabIndex        =   31
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtgrndtot 
         Height          =   300
         Left            =   2400
         TabIndex        =   28
         Top             =   3240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtdisgvn 
         Height          =   300
         Left            =   2400
         TabIndex        =   26
         Top             =   3720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtpayable 
         Height          =   300
         Left            =   2400
         TabIndex        =   24
         Top             =   4320
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtServiceName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   4
         Top             =   960
         Width           =   2415
      End
      Begin VB.ComboBox cmbServiceID 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtrpu 
         Height          =   285
         Left            =   2400
         TabIndex        =   2
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   0
         X2              =   5040
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grand Total"
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
         Height          =   195
         Left            =   480
         TabIndex        =   29
         Top             =   3240
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Given"
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
         Height          =   195
         Left            =   480
         TabIndex        =   27
         Top             =   3720
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount"
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
         Height          =   195
         Left            =   480
         TabIndex        =   25
         Top             =   4320
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service Charge"
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
         Height          =   195
         Left            =   480
         TabIndex        =   7
         Top             =   1560
         Width           =   1485
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service Name"
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
         Height          =   195
         Left            =   480
         TabIndex        =   6
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service ID"
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
         Height          =   195
         Left            =   480
         TabIndex        =   5
         Top             =   360
         Width           =   1020
      End
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   375
      Left            =   10080
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
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
      CurrentDate     =   38353
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   360
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      FormHeightDT    =   9300
      FormWidthDT     =   12465
      FormScaleHeightDT=   8790
      FormScaleWidthDT=   12345
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   615
      Left            =   4440
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   7455
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IN PATIENTS HOSPITAL SERVICES"
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
      Left            =   3000
      TabIndex        =   13
      Top             =   360
      Width           =   6945
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
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
      Height          =   195
      Left            =   9120
      TabIndex        =   12
      Top             =   1560
      Width           =   810
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No"
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
      Height          =   195
      Left            =   4920
      TabIndex        =   11
      Top             =   1440
      Width           =   615
   End
End
Attribute VB_Name = "frmInPatientServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub cmbInPatientID_Click()
Dim rsSelPatient As Recordset
Set rsSelPatient = New ADODB.Recordset
Dim rsSelAdmit As Recordset
Set rsSelAdmit = New ADODB.Recordset

rsSelPatient.Open " Select * from In_Patient_Details where Patient_ID = '" & cmbInPatientID & "'", cnPatients, adOpenDynamic, adLockReadOnly
rsSelAdmit.Open "Select * from Admission_Details where Patient_ID = '" & cmbInPatientID & "'", cnPatients, adOpenDynamic, adLockReadOnly
cmbAdmitID.clear

If rsSelPatient.RecordCount = 1 Then
    txtPatientName = rsSelPatient(1) & " " & rsSelPatient(2)
    While rsSelAdmit.EOF = False
        cmbAdmitID.AddItem rsSelAdmit(0)
        rsSelAdmit.MoveNext
    Wend
    
Else
    MsgBox "An Error Occured"
    rsSelPatient.Close
    rsSelAdmit.Close
    Exit Sub
End If
rsSelPatient.Close
rsSelAdmit.Close


End Sub

Private Sub cmbServiceID_Click()
Dim rsAddSerName As Recordset
Set rsAddSerName = New ADODB.Recordset


rsAddSerName.Open "Select * from Services where Channel_Service_ID = '" & cmbServiceID & "'", cnPatients, adOpenDynamic, adLockReadOnly

If rsAddSerName.RecordCount > 1 Then
    MsgBox " Database Error"
    Exit Sub
ElseIf rsAddSerName.RecordCount = 0 Then
    txtmedname = ""
    txtRPU = "0."
    
Else
    txtServiceName = rsAddSerName(1)
    txtRPU = rsAddSerName(3)
    txtgrndtot = txtRPU
    txtdisgvn = "0"
    txtpayable = Val(txtRPU) - Val(txtdisgvn)

    
End If

rsAddSerName.Close



End Sub

Private Sub cmdSave_Click()

Dim rsChkPatient As Recordset
Set rsChkPatient = New ADODB.Recordset

rsChkPatient.Open "select * from In_Patient_Discharge where Admission_ID = '" & cmbAdmitID & "'", cnPatients, adOpenDynamic, adLockReadOnly
If rsChkPatient.EOF = False Then
    MsgBox "The Patient has been already discharged", vbCritical
    Exit Sub
End If
rsChkPatient.Close

If cmbAdmitID = "" Then
    MsgBox "Please enter the Admssion ID", vbCritical, "Error Occured"
    Exit Sub
End If



Dim rsAddSer As Recordset
Set rsAddSer = New ADODB.Recordset

rsAddSer.Open "select * from InPatient_Services", cnPatients, adOpenDynamic, adLockPessimistic


If MsgBox("Are you sure you want to add the record to the database?", vbQuestion + vbYesNo) = vbYes Then
rsAddSer.AddNew
rsAddSer(0) = txtBillID
rsAddSer(1) = cmbInPatientID
rsAddSer(2) = cmbAdmitID
rsAddSer(3) = cmbServiceID
rsAddSer(4) = Format(DTPDate, "short date")
rsAddSer(5) = Format(DTPSDate, "short date")
rsAddSer(6) = Format(DTPSTime, "short Time")
rsAddSer(7) = txtgrndtot
rsAddSer(8) = txtdisgvn
rsAddSer(9) = txtpayable


rsAddSer.Update
rsAddSer.Close
Form_Load
Exit Sub
End If
rsAddSer.Close

End Sub

Private Sub cmdViewAdmission_Click()
frmDisplayAdmissionDetails.Show
End Sub

Private Sub cmdViewPatient_Click()
frmDisplayInPatient.Show
End Sub

Private Sub cmdViewService_Click()
frmService.Show
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Call Functions.DisableMenu
End Sub

Private Sub Form_Deactivate()
Call Functions.EnableMenu
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 22 Then KeyAscii = 0: Exit Sub
    KeyAscii = DataEntryValidation(KeyAscii, ActiveControl.Tag)
End Sub

Private Sub Form_Load()
Call Functions.DisableMenu
Me.WindowState = vbMaximized

Call AddInPatientDetails
Call AddServiceDetails
Call GenerateBillID

DTPDate = Date
DTPSDate = Date
DTPSTime = Time

End Sub


Public Sub AddInPatientDetails()
Dim rsAddPatient As Recordset
Set rsAddPatient = New ADODB.Recordset

rsAddPatient.Open "Select * from In_Patient_Details", cnPatients, adOpenDynamic, adLockReadOnly

If rsAddPatient.EOF = False Then
rsAddPatient.MoveFirst

While rsAddPatient.EOF = False
    cmbInPatientID.AddItem rsAddPatient(0)
    cmbInPatientID.Text = rsAddPatient(0)
    rsAddPatient.MoveNext
Wend


End If

rsAddPatient.Close



End Sub

Public Sub AddServiceDetails()
Dim rsAddSer As Recordset
Set rsAddSer = New ADODB.Recordset

rsAddSer.Open "Select * from Services", cnPatients, adOpenDynamic, adLockReadOnly

If rsAddSer.EOF = False Then
rsAddSer.MoveFirst

While rsAddSer.EOF = False
    cmbServiceID.AddItem rsAddSer(0)
    cmbServiceID.Text = rsAddSer(0)
    rsAddSer.MoveNext
Wend

End If

rsAddSer.Close





End Sub

Public Sub GenerateBillID()

    Dim rsAddPatient As Recordset
    Dim MID As String
    Set rsAddPatient = New ADODB.Recordset
  
    MID = Functions.UID(6, "ISerID_")
    rsAddPatient.Open " Select * from InPatient_Services", cnPatients, adOpenDynamic, adLockReadOnly
    While rsAddPatient.EOF = False
        If rsAddPatient(0) = MID Then
            MID = Functions.UID(6, "ISerID_")
            rsAddPatient.MoveFirst
        End If
    rsAddPatient.MoveNext
    Wend
    rsAddPatient.Close
    txtBillID = MID


End Sub

Private Sub Label12_Click()

End Sub

Private Sub txtdis_Change()
txttotamt = Val(txtAmount) - Val(txtdis)
End Sub

Private Sub txtqty_Change()

txtAmount = Val(txtRPU) * Val(txtqty)

End Sub


Private Sub Form_Unload(Cancel As Integer)
Call Functions.EnableMenu
End Sub

