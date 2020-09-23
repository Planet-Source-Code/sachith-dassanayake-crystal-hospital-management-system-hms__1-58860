VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmIPDischarge 
   BackColor       =   &H00FF8080&
   Caption         =   "In Patient Discharge Details"
   ClientHeight    =   6630
   ClientLeft      =   1110
   ClientTop       =   450
   ClientWidth     =   10395
   Icon            =   "frmIPDischarge.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   10395
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   28
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   300
      Left            =   600
      TabIndex        =   27
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdIPViewBill 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Patient &Bill"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8760
      Picture         =   "frmIPDischarge.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Discharge Details"
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
      Height          =   3015
      Left            =   240
      TabIndex        =   13
      Top             =   840
      Width           =   7695
      Begin VB.TextBox txtFields 
         DataField       =   "Discharge_Time"
         Height          =   405
         Index           =   3
         Left            =   3480
         TabIndex        =   20
         Top             =   2040
         Width           =   2655
      End
      Begin VB.CommandButton cmdAdmissionID 
         Caption         =   "..."
         Height          =   255
         Left            =   6240
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cmbAdmissionID 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Discharge_ID"
         Height          =   285
         Index           =   0
         Left            =   3480
         TabIndex        =   14
         Top             =   480
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker DTPTime1 
         Height          =   375
         Left            =   3480
         TabIndex        =   17
         Top             =   2040
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         Format          =   45678594
         CurrentDate     =   38367
      End
      Begin MSComCtl2.DTPicker DTPDate1 
         Height          =   375
         Left            =   3480
         TabIndex        =   18
         Top             =   1440
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
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
         Format          =   45678593
         CurrentDate     =   38367
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Discharge_Date"
         Height          =   285
         Index           =   2
         Left            =   3480
         TabIndex        =   19
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Admission_ID"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3480
         TabIndex        =   21
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Discharge Time:"
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
         Left            =   1440
         TabIndex        =   25
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Discharge Date:"
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
         Left            =   1440
         TabIndex        =   24
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Admission ID:"
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
         Left            =   1440
         TabIndex        =   23
         Top             =   975
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Discharge ID:"
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
         Left            =   1440
         TabIndex        =   22
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
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
      Left            =   240
      TabIndex        =   7
      Top             =   4080
      Width           =   7695
      Begin VB.CommandButton cmdLast 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   6720
         Picture         =   "frmIPDischarge.frx":5DB4
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   6000
         Picture         =   "frmIPDischarge.frx":6289
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   1065
         Picture         =   "frmIPDischarge.frx":6764
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdFirst 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   360
         Picture         =   "frmIPDischarge.frx":6C45
         Style           =   1  'Graphical
         TabIndex        =   8
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
         TabIndex        =   12
         Top             =   480
         Width           =   3960
      End
   End
   Begin VB.Frame Frame3 
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
      Height          =   4455
      Left            =   8160
      TabIndex        =   0
      Top             =   840
      Width           =   2055
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   480
         Picture         =   "frmIPDischarge.frx":711B
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   480
         Picture         =   "frmIPDischarge.frx":75CA
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   480
         Picture         =   "frmIPDischarge.frx":7A70
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Update"
         Height          =   780
         Left            =   480
         Picture         =   "frmIPDischarge.frx":7F7C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   480
         Picture         =   "frmIPDischarge.frx":842A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdViewAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&View"
         Height          =   780
         Left            =   480
         Picture         =   "frmIPDischarge.frx":892E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1320
         Width           =   1095
      End
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
      FormHeightDT    =   7140
      FormWidthDT     =   10515
      FormScaleHeightDT=   6630
      FormScaleWidthDT=   10395
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "DISCHARGE DETAILS"
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
      Index           =   4
      Left            =   3000
      TabIndex        =   29
      Top             =   240
      Width           =   4260
   End
End
Attribute VB_Name = "frmIPDischarge"
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
Public strAdmitID As String
Public strPatID As String
Public FromDischarge As Boolean

Private Sub cmdIPViewBill_Click()
Dim rsPatID As Recordset
Set rsPatID = New ADODB.Recordset

rsPatID.Open "select * from Admission_Details where Admission_ID = '" & txtFields(1) & "'", cnPatients, adOpenDynamic, adLockReadOnly
If rsPatID.EOF = True Then
    MsgBox "Admission ID not found", vbCritical
    rsPatID.Close
    Exit Sub
End If

strPatID = rsPatID(1)
rsPatID.Close
    
strAdmitID = txtFields(1)
FromDischarge = True
frmIPBill.Show
Unload Me

End Sub

Private Sub cmdViewAll_Click()
frmDisplayIPDischarge.Show
End Sub

Private Sub Form_Load()
  
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select Discharge_ID,Admission_ID,Discharge_Date,Discharge_Time from In_Patient_Discharge", cnPatients, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
    oText.Locked = True
  Next
  
  DTPDate1.Value = Date
  DTPTime1.Value = Time
  
 
  cmbAdmissionID.clear
  Call addAdmssionID
  FromDischarge = False

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
  
     Dim rsDisID As Recordset
    Dim PID As String
    Set rsDisID = New ADODB.Recordset
  
    PID = Functions.UID(6, "DisID_")
    rsDisID.Open " Select * from In_Patient_Discharge", cnPatients, adOpenKeyset, adLockPessimistic
    While rsDisID.EOF = False
        If rsDisID(0) = PID Then
            ID = True
            PID = Functions.UID(6, "DisID_")
            rsDisID.MoveFirst
        Else
            ID = False
        End If
    rsDisID.MoveNext
    Wend
    rsDisID.Close
    
 
  cmbAdmissionID.clear
  Call addAdmssionID
  
  
  
  
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    txtFields(0) = PID
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

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  
    
 
  cmbAdmissionID.clear
  Call addAdmssionID
  
  
  cmbAdmissionID = txtFields(1)
  DTPDate1.Value = txtFields(2)
  DTPTime1.Value = txtFields(3)
 
  
  
  
  
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
  
  txtFields(1) = cmbAdmissionID
  txtFields(2) = Format(DTPDate1.Value, "short Date")
  txtFields(3) = Format(DTPTime1.Value, "short Time")
  
  If txtFields(0) = "" Then
    MsgBox "Invalid Discharge ID", vbCritical
    Exit Sub
  End If
  
  If txtFields(1) = "" Then
    MsgBox "Please enter a valid Admission ID", vbCritical
    Exit Sub
  End If
  
  
  Dim rsChk As Recordset
  Set rsChk = New ADODB.Recordset
  
  rsChk.Open "Select * from In_Patient_Discharge where Admission_ID = '" & cmbAdmissionID & "'", cnPatients, adOpenDynamic, adLockReadOnly
  If rsChk.EOF = False Then
    MsgBox "The Patient has alredy Discharged", vbCritical
    rsChk.Close
    Exit Sub
  End If
  rsChk.Close
  
   
  
  If MsgBox("Are you sure you want to disharge the patient?", vbQuestion + vbYesNo, "Confirm Patient Discharge") = vbNo Then
    Exit Sub
  End If
  
  Dim rsBed As Recordset
  Set rsBed = New ADODB.Recordset
  
  rsBed.Open "Select * from Bed_Details where Admission_ID = '" & cmbAdmissionID & "'", cnPatients, adOpenDynamic, adLockPessimistic
  
  If rsBed.EOF = True Then
    MsgBox "A Bed was not assigned to this patient", vbCritical
  ElseIf rsBed.RecordCount = 1 Then
    
    rsBed(2) = True
    rsBed(3) = ""
    rsBed.Update
  Else
    MsgBox "An Error Occured", vbCritical
    Exit Sub
  End If
  rsBed.Close
  
  
  
  

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
  'cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  'cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
  
  
  cmdViewAll.Visible = bVal
  
  DTPDate1.Visible = Not bVal
  DTPTime1.Visible = Not bVal
  
  cmdAdmissionID.Visible = Not bVal
  cmbAdmissionID.Visible = Not bVal
  
  
End Sub

Private Sub addAdmssionID()

Dim rsAddAdmission As Recordset
Set rsAddAdmission = New ADODB.Recordset

rsAddAdmission.Open "select * from Admission_Details", cnPatients, adOpenDynamic, adLockReadOnly


While rsAddAdmission.EOF = False
    cmbAdmissionID.AddItem rsAddAdmission(0)
    rsAddAdmission.MoveNext

Wend

rsAddAdmission.Close

End Sub
