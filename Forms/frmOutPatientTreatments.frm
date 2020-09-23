VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmOutPatientTreatments 
   BackColor       =   &H00FF8080&
   Caption         =   "OutPatient Treatments Details"
   ClientHeight    =   9165
   ClientLeft      =   1110
   ClientTop       =   450
   ClientWidth     =   8145
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9165
   ScaleWidth      =   8145
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbPatientID 
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   1080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ComboBox cmbDoctorID 
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   1680
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdDoctorID 
      Caption         =   "..."
      Height          =   255
      Left            =   5640
      TabIndex        =   31
      Top             =   1680
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSComCtl2.DTPicker DTPTime1 
      Height          =   375
      Left            =   3000
      TabIndex        =   30
      Top             =   2880
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Format          =   46137346
      CurrentDate     =   38376
   End
   Begin MSComCtl2.DTPicker DTPDate1 
      Height          =   375
      Left            =   3000
      TabIndex        =   29
      Top             =   2280
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      Format          =   46137345
      CurrentDate     =   38376
   End
   Begin VB.CommandButton cmdPatientID 
      Caption         =   "..."
      Height          =   255
      Left            =   5640
      TabIndex        =   28
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   7560
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      FormHeightDT    =   9675
      FormWidthDT     =   8265
      FormScaleHeightDT=   9165
      FormScaleWidthDT=   8145
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
      Height          =   1455
      Left            =   480
      TabIndex        =   20
      Top             =   7560
      Width           =   7215
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         Height          =   780
         Left            =   480
         Picture         =   "frmOutPatientTreatments.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         Height          =   780
         Left            =   1800
         Picture         =   "frmOutPatientTreatments.frx":04AF
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   780
         Left            =   3120
         Picture         =   "frmOutPatientTreatments.frx":0954
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   780
         Left            =   4440
         Picture         =   "frmOutPatientTreatments.frx":0E0D
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   780
         Left            =   5760
         Picture         =   "frmOutPatientTreatments.frx":12B3
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Update"
         Height          =   780
         Left            =   480
         Picture         =   "frmOutPatientTreatments.frx":17BF
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   780
         Left            =   1800
         Picture         =   "frmOutPatientTreatments.frx":1C6D
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
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
      Left            =   480
      TabIndex        =   14
      Top             =   6120
      Width           =   7215
      Begin VB.CommandButton cmdLast 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   6240
         Picture         =   "frmOutPatientTreatments.frx":2171
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   5520
         Picture         =   "frmOutPatientTreatments.frx":2646
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   1065
         Picture         =   "frmOutPatientTreatments.frx":2B21
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdFirst 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   360
         Picture         =   "frmOutPatientTreatments.frx":3002
         Style           =   1  'Graphical
         TabIndex        =   15
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
         TabIndex        =   19
         Top             =   480
         Width           =   3480
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Prescription"
      Height          =   1005
      Index           =   6
      Left            =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   4680
      Width           =   2535
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Description"
      Height          =   870
      Index           =   5
      Left            =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   3465
      Width           =   2535
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Time"
      Height          =   285
      Index           =   4
      Left            =   3000
      TabIndex        =   9
      Top             =   2880
      Width           =   2535
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Date"
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   7
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Doctor_ID"
      Height          =   285
      Index           =   2
      Left            =   3000
      TabIndex        =   5
      Top             =   1665
      Width           =   2535
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Patient_ID"
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtFields 
      DataField       =   "OPHistoryID"
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   1
      Top             =   540
      Width           =   2535
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Prescription:"
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
      Index           =   6
      Left            =   840
      TabIndex        =   12
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Description:"
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
      Index           =   5
      Left            =   840
      TabIndex        =   10
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Time:"
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
      Index           =   4
      Left            =   840
      TabIndex        =   8
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Date:"
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
      Index           =   3
      Left            =   840
      TabIndex        =   6
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Doctor ID:"
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
      Index           =   2
      Left            =   840
      TabIndex        =   4
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Patient ID:"
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
      Left            =   840
      TabIndex        =   2
      Top             =   1095
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Treatment ID:"
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
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "frmOutPatientTreatments"
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

Private Sub cmdDoctorID_Click()
frmDoctorDetails.Show
End Sub

Private Sub cmdPatientID_Click()
frmDisplayOutPatient.Show
End Sub

Private Sub Form_Activate()
Call Functions.DisableMenu
End Sub

Private Sub Form_Deactivate()
Call Functions.EnableMenu
End Sub

Private Sub Form_Load()
 
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select OPHistoryID,Patient_ID,Doctor_ID,Date,Time,Description,Prescription from OutPatient_Treatments", cnPatients, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
    oText.Locked = True
  Next
  
  
  Dim rsAddPatientID As Recordset
  Set rsAddPatientID = New ADODB.Recordset
  
  Dim rsAddDoctorID As Recordset
  Set rsAddDoctorID = New ADODB.Recordset
  
  rsAddPatientID.Open "Select * from Patient_Details", cnPatients, adOpenDynamic, adLockReadOnly
  
  While rsAddPatientID.EOF = False
  cmbPatientID.AddItem rsAddPatientID(0)
  rsAddPatientID.MoveNext
  Wend
  rsAddPatientID.Close
  
  rsAddDoctorID.Open "select * from Doctor_Details", cnPatients, adOpenDynamic, adLockReadOnly
  
  While rsAddDoctorID.EOF = False
  cmbDoctorID.AddItem rsAddDoctorID(0)
  
  rsAddDoctorID.MoveNext
  
  Wend
  rsAddDoctorID.Close
  
  
  
  
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

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  
    Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    oText.Locked = False
  Next
  txtFields(0).Locked = True
  
  DTPDate1 = Date
  DTPTime1 = Time
  
    Dim rsAddID As Recordset
    Dim GID As String
    Set rsAddID = New ADODB.Recordset
  
    GID = Functions.UID(6, "OPTreat_")
    rsAddID.Open " Select * from OutPatient_Treatments", cnPatients, adOpenKeyset, adLockPessimistic
    While rsAddID.EOF = False
        If rsAddID(0) = GID Then
            GID = Functions.UID(6, "OPTreat_")
            rsAddID.MoveFirst
        End If
    rsAddID.MoveNext
    Wend
    rsAddID.Close
  
  
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    txtFields(0) = GID
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

    cmbPatientID = txtFields(1)
    cmbDoctorID = txtFields(2)
    DTPDate1 = txtFields(3)
    DTPTime1 = txtFields(4)


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
  
 
  
  
  txtFields(1) = cmbPatientID
  txtFields(2) = cmbDoctorID
  txtFields(3) = DTPDate1
  txtFields(4) = DTPTime1
  
   If txtFields(1) = "" Then
    MsgBox "Please Select Patient ID", vbCritical
    Exit Sub
  End If
  
  If txtFields(2) = "" Then
    MsgBox "Please Select Doctor ID", vbCritical
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
  
      Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    oText.Locked = True
  Next

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
  
  cmdDoctorID.Visible = Not bVal
  cmdPatientID.Visible = Not bVal
  DTPDate1.Visible = Not bVal
  DTPTime1.Visible = Not bVal
  
  cmbDoctorID.Visible = Not bVal
  cmbPatientID.Visible = Not bVal
  
  
End Sub

