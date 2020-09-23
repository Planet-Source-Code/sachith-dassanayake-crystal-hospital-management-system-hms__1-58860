VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmAddDoctorVisits 
   BackColor       =   &H00FF8080&
   Caption         =   "Visit Details"
   ClientHeight    =   9240
   ClientLeft      =   1275
   ClientTop       =   1455
   ClientWidth     =   8835
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddDoctorVisits.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   8835
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   7575
      Left            =   480
      TabIndex        =   31
      Top             =   840
      Visible         =   0   'False
      Width           =   7935
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   975
         Left            =   4440
         Picture         =   "frmAddDoctorVisits.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   6480
         Width           =   1095
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&OK"
         Height          =   975
         Left            =   2760
         Picture         =   "frmAddDoctorVisits.frx":5C2E
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   6480
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5055
         Left            =   600
         TabIndex        =   32
         Top             =   600
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   8916
         _Version        =   393216
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         HighLight       =   2
         SelectionMode   =   1
      End
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Add"
      Height          =   780
      Left            =   1680
      Picture         =   "frmAddDoctorVisits.frx":6104
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Edit"
      Height          =   780
      Left            =   2880
      Picture         =   "frmAddDoctorVisits.frx":65B3
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdPatient 
      Caption         =   "..."
      Height          =   315
      Left            =   6720
      TabIndex        =   7
      Top             =   3360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdDoc 
      Caption         =   "..."
      Height          =   315
      Left            =   6720
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ComboBox cmbPatientID 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox cmbDoctorID 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   2760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker DTPTime1 
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   2160
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Format          =   46006274
      CurrentDate     =   38355
   End
   Begin MSComCtl2.DTPicker DTPDate1 
      Height          =   375
      Left            =   4200
      TabIndex        =   22
      Top             =   1560
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Format          =   46006273
      CurrentDate     =   38355
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Description"
      Height          =   990
      Index           =   5
      Left            =   4200
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PAtient_ID"
      Height          =   285
      Index           =   4
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Doctor_ID"
      Height          =   285
      Index           =   3
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Visit_Time"
      Height          =   285
      Index           =   2
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2145
      Width           =   2415
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Visit_Date"
      Height          =   285
      Index           =   1
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Visit_ID"
      Height          =   285
      Index           =   0
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1020
      Width           =   2415
   End
   Begin VB.ComboBox cmbAdmitID 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   3840
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Admission_ID"
      Height          =   285
      Index           =   6
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   780
      Left            =   2880
      Picture         =   "frmAddDoctorVisits.frx":6A58
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   7080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Update"
      Height          =   780
      Left            =   1680
      Picture         =   "frmAddDoctorVisits.frx":6F5C
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   780
      Left            =   6600
      Picture         =   "frmAddDoctorVisits.frx":740A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      Height          =   780
      Left            =   5280
      Picture         =   "frmAddDoctorVisits.frx":7916
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
      Height          =   780
      Left            =   4080
      Picture         =   "frmAddDoctorVisits.frx":7DBC
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   1440
      Picture         =   "frmAddDoctorVisits.frx":8275
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   585
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   2145
      Picture         =   "frmAddDoctorVisits.frx":874B
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   585
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   6600
      Picture         =   "frmAddDoctorVisits.frx":8C2C
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   585
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   7320
      Picture         =   "frmAddDoctorVisits.frx":9107
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   585
   End
   Begin VB.CommandButton cmdAdmission 
      Caption         =   "..."
      Height          =   315
      Left            =   6720
      TabIndex        =   9
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
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
      FormHeightDT    =   9750
      FormWidthDT     =   8955
      FormScaleHeightDT=   9240
      FormScaleWidthDT=   8835
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Left            =   2880
      TabIndex        =   13
      Top             =   6120
      Width           =   3480
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   855
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Width           =   7335
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   1095
      Left            =   1440
      Shape           =   4  'Rounded Rectangle
      Top             =   6960
      Width           =   6495
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   4695
      Left            =   1200
      Shape           =   4  'Rounded Rectangle
      Top             =   840
      Width           =   6735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Doctor Visits"
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
      Left            =   3360
      TabIndex        =   37
      Top             =   240
      Width           =   2550
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Admission ID:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   2280
      TabIndex        =   36
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Description:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   30
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Patient ID:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   29
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Doctor ID:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   2280
      TabIndex        =   28
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Visit Time:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   27
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Visit Date:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   26
      Top             =   1575
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Visit ID:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   25
      Top             =   1020
      Width           =   1815
   End
End
Attribute VB_Name = "frmAddDoctorVisits"
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



Private Sub cmbPatientID_Click()
Dim rsAddAdmitID As Recordset
Set rsAddAdmitID = New ADODB.Recordset
cmbAdmitID.clear

rsAddAdmitID.Open "select * from Admission_Details where Patient_ID='" & cmbPatientID & "'", cnPatients, adOpenDynamic, adLockReadOnly

While rsAddAdmitID.EOF = False
cmbAdmitID.AddItem rsAddAdmitID(0)
rsAddAdmitID.MoveNext
Wend

rsAddAdmitID.Close



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
  adoPrimaryRS.Open "select Visit_ID,Visit_Date,Visit_Time,Doctor_ID,Patient_ID,Admission_ID,Description from Visit_Details", cnPatients, adOpenDynamic, adLockPessimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next

    DTPDate1 = Date
    DTPTime1 = Time
    Call addPatientID
    Call addDoctorID
    
  mbDataChanged = False
End Sub

Private Sub Form_Resize()
  On Error Resume Next
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
  lblStatus.Caption = "Doctor Visit Record: " & CStr(adoPrimaryRS.AbsolutePosition)
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
  On Error GoTo AddErr
    
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
  If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "Delete Record") = vbNo Then
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
  adoPrimaryRS.Requery
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

    
    txtFields(1) = DTPDate1.Value
    txtFields(2) = DTPTime1.Value
    txtFields(3) = cmbDoctorID
    txtFields(4) = cmbPatientID
    txtFields(6) = cmbAdmitID

  If txtFields(3) = "" Or txtFields(4) = "" Or txtFields(6) = "" Then
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
  txtFields(5).Locked = bVal
  DTPDate1.Visible = Not bVal
  DTPTime1.Visible = Not bVal
  cmbDoctorID.Visible = Not bVal
  cmbPatientID.Visible = Not bVal
  cmbAdmitID.Visible = Not bVal
  cmdAdmission.Visible = Not bVal
  cmdDoc.Visible = Not bVal
  cmdPatient.Visible = Not bVal
  
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


