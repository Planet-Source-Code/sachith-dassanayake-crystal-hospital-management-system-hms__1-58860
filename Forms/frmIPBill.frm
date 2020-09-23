VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmIPBill 
   BackColor       =   &H00FF8080&
   Caption         =   "In Patient Billing "
   ClientHeight    =   10035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11595
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmIPBill.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10035
   ScaleWidth      =   11595
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "Duration"
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
      Left            =   6360
      TabIndex        =   41
      Top             =   1200
      Width           =   4815
      Begin MSComCtl2.DTPicker DTPAdmit 
         Height          =   375
         Left            =   2520
         TabIndex        =   42
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   47644673
         CurrentDate     =   38353
      End
      Begin MSComCtl2.DTPicker DTPDisDate 
         Height          =   375
         Left            =   2520
         TabIndex        =   43
         Top             =   1560
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Format          =   47644673
         CurrentDate     =   38357
      End
      Begin VB.TextBox txtAdmitDate 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         TabIndex        =   44
         Top             =   480
         Width           =   2000
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "Date of Admission"
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
         Left            =   240
         TabIndex        =   48
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF8080&
         Caption         =   "Date"
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
         Left            =   240
         TabIndex        =   47
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label17 
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
         Left            =   240
         TabIndex        =   46
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label18 
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
         Left            =   240
         TabIndex        =   45
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Charges"
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
      Height          =   3735
      Left            =   480
      TabIndex        =   15
      Top             =   5040
      Width           =   10575
      Begin VB.TextBox txtVat 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7920
         TabIndex        =   34
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtTotal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7920
         TabIndex        =   33
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txtDiscount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7920
         TabIndex        =   32
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtNetValue 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7920
         TabIndex        =   31
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox txtVisitCharges 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   25
         Top             =   480
         Width           =   2000
      End
      Begin VB.TextBox txtService 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   24
         Top             =   1080
         Width           =   2000
      End
      Begin VB.TextBox txtMedCharges 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   23
         Top             =   1680
         Width           =   2000
      End
      Begin VB.TextBox txtRoomCharges 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   22
         Top             =   2280
         Width           =   2000
      End
      Begin VB.TextBox txtHospitalCharges 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2640
         TabIndex        =   21
         Top             =   2880
         Width           =   2000
      End
      Begin VB.CommandButton cmdDocVisit 
         Caption         =   "..."
         Height          =   255
         Left            =   4800
         TabIndex        =   20
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton cmdServiceCharges 
         Caption         =   "..."
         Height          =   255
         Left            =   4800
         TabIndex        =   19
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmdMedicine 
         Caption         =   "..."
         Height          =   255
         Left            =   4800
         TabIndex        =   18
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         Height          =   255
         Left            =   4800
         TabIndex        =   17
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "..."
         Height          =   255
         Left            =   4800
         TabIndex        =   16
         Top             =   2880
         Width           =   495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FF8080&
         Caption         =   "VAT"
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
         Left            =   6120
         TabIndex        =   40
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FF8080&
         Caption         =   "Total"
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
         Left            =   6120
         TabIndex        =   39
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FF8080&
         Caption         =   "Discount"
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
         Left            =   6120
         TabIndex        =   38
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF8080&
         Caption         =   "Net Value"
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
         Left            =   6120
         TabIndex        =   37
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FF8080&
         Caption         =   "%"
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
         Left            =   9960
         TabIndex        =   36
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FF8080&
         Caption         =   "%"
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
         Left            =   9960
         TabIndex        =   35
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF8080&
         Caption         =   "Doctor Visit Charges"
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
         TabIndex        =   30
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF8080&
         Caption         =   "Hospital Service Charges"
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
         TabIndex        =   29
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF8080&
         Caption         =   "Medicine Charges"
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
         TabIndex        =   28
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FF8080&
         Caption         =   "Room Charges"
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
         TabIndex        =   27
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FF8080&
         Caption         =   "Hospital Charges"
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
         TabIndex        =   26
         Top             =   2880
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   3495
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   5415
      Begin VB.ComboBox cmbAdmissionID 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox cmbPatient 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
      Begin VB.TextBox txtRoomWardID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Top             =   2280
         Width           =   2000
      End
      Begin VB.TextBox txtBedID 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   6
         Top             =   2880
         Width           =   2000
      End
      Begin VB.TextBox txtStatus 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   5
         Top             =   1680
         Width           =   2000
      End
      Begin VB.CommandButton Command6 
         Caption         =   "..."
         Height          =   255
         Left            =   4680
         TabIndex        =   4
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton Command7 
         Caption         =   "..."
         Height          =   255
         Left            =   4680
         TabIndex        =   3
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "Patient ID"
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
         TabIndex        =   14
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF8080&
         Caption         =   "Room / Ward ID"
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
         TabIndex        =   13
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF8080&
         Caption         =   "Bed ID"
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
         TabIndex        =   12
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FF8080&
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
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FF8080&
         Caption         =   "Status"
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
         TabIndex        =   10
         Top             =   1680
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   855
      Left            =   7800
      Picture         =   "frmIPBill.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8880
      Width           =   1335
   End
   Begin VB.CommandButton cmdPayBill 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      Height          =   855
      Left            =   3600
      Picture         =   "frmIPBill.frx":5C2E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8880
      Width           =   1335
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
      FormHeightDT    =   10545
      FormWidthDT     =   11715
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   11595
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "IN PATIENT BILLING"
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
      Left            =   3840
      TabIndex        =   49
      Top             =   240
      Width           =   4200
   End
End
Attribute VB_Name = "frmIPBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RoomCharge As Double
Private Sub cmbAdmissionID_Click()
Dim ctl As Control
For Each ctl In Controls
 If TypeOf ctl Is TextBox Then
    ctl.Text = ""
    ctl.Locked = True
    ctl.Alignment = 1
 End If
Next
txtStatus.Alignment = 0
txtRoomWardID.Alignment = 0
txtBedID.Alignment = 0

Dim rsCheckStatus As Recordset
Dim rsAddData As Recordset

Set rsCheckStatus = New ADODB.Recordset
Set rsAddData = New ADODB.Recordset
DTPDisDate.Enabled = True

rsCheckStatus.Open "Select * from In_Patient_Discharge where Admission_ID = '" & cmbAdmissionID & " ' ", cnPatients, adOpenDynamic, adLockReadOnly
If rsCheckStatus.EOF = False Then
    txtStatus = "Discharged"
    DTPDisDate = rsCheckStatus(2)
    DTPDisDate.Enabled = False
Else
    txtStatus = "Under Treatements"
End If

rsCheckStatus.Close

rsAddData.Open "Select * from Admission_Details where Admission_ID = '" & cmbAdmissionID & "'", cnPatients, adOpenDynamic, adLockReadOnly

If rsAddData.EOF = False Then
    txtAdmitDate = Format(rsAddData(6), "short date")
    DTPAdmit = Format(txtAdmitDate, "short date")
    txtRoomWardID = rsAddData(3)
    txtBedID = rsAddData(4)
End If
rsAddData.Close
Call FillData
Call MedicineCharges
Call ServiceCharges
Call HospitalCharges
DTPDisDate_Change
End Sub



Private Sub cmbPatient_Click()
    cmbAdmissionID.clear
    Dim ctl As Control
    For Each ctl In Controls
         If TypeOf ctl Is TextBox Then
           ctl.Text = ""
        End If
Next
    addAdmissionID cmbPatient
    
End Sub






Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdDocVisit_Click()
frmDoctorVisit.Show
End Sub

Private Sub cmdMedicine_Click()
frmInPatientOrders.Show
End Sub

Private Sub cmdPayBill_Click()
Dim ct As Control

For Each ctl In Controls
    If TypeOf ctl Is TextBox Then
        If ctl.Text = "" Then
            MsgBox "one of Required Field is missing", vbCritical
            ctl.SetFocus
            Exit Sub
        End If
    End If
Next





If Trim(txtStatus) <> "Discharged" Then
    MsgBox "You Cannot save the bill details until patient discharge", vbInformation
    Exit Sub
End If

If MsgBox("Do you want to save the record and view patient bill?", vbQuestion + vbYesNo) = vbNo Then
    Exit Sub
End If


Dim rsAdmit As Recordset
Set rsAdmit = New ADODB.Recordset

rsAdmit.Open "select * from Patient_Bill where Admission_ID = '" & cmbAdmissionID & "'", cnPatients, adOpenDynamic, adLockReadOnly
If rsAdmit.EOF = False Then
    MsgBox "The information has already stored in the database." & vbCrLf & "You can not add the same record again", vbCritical, "Error Occured"
    rsAdmit.Close
    Exit Sub
End If
rsAdmit.Close


Call SaveData

End Sub

Private Sub cmdServiceCharges_Click()
frmIPServiceDetails.Show
End Sub

Private Sub Command3_Click()

End Sub

Private Sub DTPDisDate_Change()
Dim tot As Double
If txtAdmitDate <> "" Then
 txtRoomCharges = RoomCharge * DateDiff("d", txtAdmitDate, DTPDisDate)
 

tot = Val(txtVisitCharges) + Val(txtService) + Val(txtMedCharges) + Val(txtRoomCharges) + Val(txtHospitalCharges)
txtTotal = tot + (Val(txtVat) * tot / 100)
txtNetValue = Format(txtTotal - (Val(txtTotal) * Val(txtDiscount)), "Standard")
 
End If
End Sub

Private Sub Form_Load()

If frmIPDischarge.FromDischarge = True Then
     Call addPatientID
     
     cmbPatient = frmIPDischarge.strPatID
     cmbPatient_Click
     
     cmbAdmissionID = frmIPDischarge.strAdmitID
     cmbAdmissionID_Click
    
    frmIPDischarge.FromDischarge = False
    Exit Sub
End If


Call addPatientID
DTPDisDate = Date
Call Functions.DisableMenu
End Sub

Private Sub addPatientID()

Dim rsAddPatient As Recordset
Set rsAddPatient = New ADODB.Recordset

rsAddPatient.Open "Select * from In_Patient_Details", cnPatients, adOpenDynamic, adLockReadOnly

While rsAddPatient.EOF = False
cmbPatient.AddItem (rsAddPatient(0))
rsAddPatient.MoveNext

Wend

rsAddPatient.Close
End Sub

Private Sub addAdmissionID(AddID As String)
Dim rsAddID As Recordset
Set rsAddID = New ADODB.Recordset

rsAddID.Open "Select * from Admission_Details where Patient_ID = '" & AddID & "'", cnPatients, adOpenDynamic, adLockReadOnly

While rsAddID.EOF = False
cmbAdmissionID.AddItem rsAddID(0)
rsAddID.MoveNext
Wend

rsAddID.Close
End Sub

Private Sub FillData()
Dim chk As Integer
Dim DocCharge As Double
Dim rsFill As Recordset
Set rsFill = New ADODB.Recordset
Dim rsDoc As Recordset
Set rsDoc = New ADODB.Recordset
Dim rsRoom As Recordset
Set rsRoom = New ADODB.Recordset


chk = 0
rsFill.Open "Select * from Visit_Details where Admission_ID = '" & cmbAdmissionID & "'", cnPatients, adOpenDynamic, adLockReadOnly

While rsFill.EOF = False
    rsDoc.Open "Select * from Doctor_Details where Doctor_ID = '" & rsFill(3) & "'", cnPatients, adOpenDynamic, adLockReadOnly
        If rsDoc.EOF = False Then
            DocCharge = DocCharge + Val(rsDoc(11))
        End If
    rsDoc.Close
    rsFill.MoveNext
Wend
rsFill.Close
txtVisitCharges = Format(DocCharge, "Standard")


rsFill.Open "Select * from Room_Details where Room_ID = '" & Trim(txtRoomWardID) & "'", cnPatients, adOpenDynamic, adLockReadOnly
    If rsFill.EOF = False Then
        rsRoom.Open "Select * from Room_Types Where Room_Type= '" & rsFill(1) & "'", cnPatients, adOpenDynamic, adLockReadOnly
            If rsRoom.EOF = False Then
                RoomCharge = Val(rsRoom(1))
                txtRoomCharges = Format(Val(rsRoom(1)) * DateDiff("d", txtAdmitDate, DTPDisDate), "Standard")
                chk = 1
            Else
                MsgBox "Database Error" & vbCrLf & "Please Contact Database Administrator", vbCritical
                Exit Sub
            End If
    End If
rsFill.Close


rsFill.Open "Select * from Ward_Details where Ward_ID = ' " & Trim(txtRoomWardID) & "'", cnPatients, adOpenDynamic, adLockReadOnly
    If rsFill.EOF = False Then
        If chk = 1 Then
            MsgBox "Data Error.. Please Contact Database Administrator", vbCritical
            Exit Sub
        End If
        RoomCharge = Val(rsFill(2))
        txtRoomCharges = Format(Val(rsFill(2)) * DateDiff("d", txtAdmitDate, DTPDisDate), "Standard")
    End If
rsFill.Close

End Sub


Private Sub MedicineCharges()
Dim amount As Double
Dim rsMed As Recordset
Set rsMed = New ADODB.Recordset
Dim rsMedDetails As Recordset
Set rsMedDetails = New ADODB.Recordset

If cmbAdmissionID = "" Then
    Exit Sub
End If


rsMed.Open "Select * from InPatient_Orders where AdmissionID = '" & cmbAdmissionID & "'", cnPatients, adOpenDynamic, adLockReadOnly
    
    While rsMed.EOF = False
        rsMedDetails.Open "Select * from InPatient_Order_Details where OrderID = '" & rsMed(0) & "'", cnPatients, adOpenDynamic, adLockReadOnly
            While rsMedDetails.EOF = False
                amount = amount + ((rsMedDetails(4) * Val(rsMedDetails(5))) - Val(rsMedDetails(6)))
                rsMedDetails.MoveNext
            Wend
        rsMed.MoveNext
        rsMedDetails.Close
    Wend

rsMed.Close


txtMedCharges = Format(amount, "Standard")

End Sub

Private Sub ServiceCharges()
Dim rsSer As Recordset
Dim amount As Double
Set rsSer = New ADODB.Recordset

If cmbAdmissionID = "" Then
    Exit Sub
End If

rsSer.Open "Select * from InPatient_Services where AdmissionID = '" & cmbAdmissionID & "'", cnPatients, adOpenDynamic, adLockReadOnly

While rsSer.EOF = False
    amount = amount + rsSer(9)
    rsSer.MoveNext
Wend

rsSer.Close

txtService = Format(amount, "Standard")

End Sub

Private Sub HospitalCharges()
Dim rsHospital As Recordset
Set rsHospital = New ADODB.Recordset
Dim tot As Double
rsHospital.Open "Select * from Hospital_Charges", cnPatients, adOpenDynamic, adLockReadOnly

If rsHospital.EOF = False Then

txtHospitalCharges = Format(rsHospital(4), "Standard")
txtVat = rsHospital(2)
tot = Val(txtVisitCharges) + Val(txtService) + Val(txtMedCharges) + Val(txtRoomCharges) + Val(txtHospitalCharges)
txtTotal = tot + (rsHospital(2) * tot / 100)

txtDiscount = rsHospital(3)

txtNetValue = Format(txtTotal - (Val(txtTotal) * rsHospital(3)), "Standard")


End If

rsHospital.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Functions.EnableMenu
End Sub


Private Sub SaveData()

    Dim rsAddPatient As Recordset
    Dim MID As String
    Set rsAddPatient = New ADODB.Recordset
  
    MID = Functions.UID(6, "IPBID_")
    rsAddPatient.Open " Select * from Patient_Bill", cnPatients, adOpenDynamic, adLockReadOnly
    While rsAddPatient.EOF = False
        If rsAddPatient(0) = MID Then
            MID = Functions.UID(6, "IPBID_")
            rsAddPatient.MoveFirst
        End If
    rsAddPatient.MoveNext
    Wend
    rsAddPatient.Close


    Dim rsAddBill As Recordset
    Set rsAddBill = New ADODB.Recordset
    
    rsAddBill.Open "Patient_Bill", cnPatients, adOpenDynamic, adLockPessimistic
    
    rsAddBill.AddNew
        rsAddBill(0) = MID
        rsAddBill(1) = cmbPatient
        rsAddBill(2) = cmbAdmissionID
        rsAddBill(3) = Format(DTPDisDate, "short Date")
        rsAddBill(4) = Val(txtVisitCharges)
        rsAddBill(5) = Val(txtMedCharges)
        rsAddBill(6) = Val(txtService)
        rsAddBill(7) = Val(txtRoomCharges)
        rsAddBill(8) = Val(txtHospitalCharges)
        rsAddBill(9) = Val(txtDiscount)
        rsAddBill(10) = Format(txtNetValue, "Currency")
        rsAddBill(11) = ""
    rsAddBill.Update

MsgBox "Record Saved Sucessfully", vbInformation

rsAddBill.Close

End Sub


