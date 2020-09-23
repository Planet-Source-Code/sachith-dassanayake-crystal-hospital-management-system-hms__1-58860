VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmInPatientBill 
   BackColor       =   &H00FF8080&
   Caption         =   "In Patient Bill Details"
   ClientHeight    =   10830
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   8730
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInPatientBill.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   10830
   ScaleWidth      =   8730
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "In Patient Details"
      ForeColor       =   &H00FFFFFF&
      Height          =   6135
      Left            =   1080
      TabIndex        =   14
      Top             =   1320
      Width           =   6735
      Begin VB.TextBox txtFields 
         DataField       =   "Patient_Bill_ID"
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   26
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Patient_ID"
         Height          =   285
         Index           =   1
         Left            =   3120
         TabIndex        =   25
         Top             =   915
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Admission_ID"
         Height          =   285
         Index           =   2
         Left            =   3120
         TabIndex        =   24
         Top             =   1365
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "DischargeDate"
         Height          =   285
         Index           =   3
         Left            =   3120
         TabIndex        =   23
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Doctor_Charges"
         Height          =   285
         Index           =   4
         Left            =   3120
         TabIndex        =   22
         Top             =   2235
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Medicine_Charges"
         Height          =   285
         Index           =   5
         Left            =   3120
         TabIndex        =   21
         Top             =   2685
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Services_Charges"
         Height          =   285
         Index           =   6
         Left            =   3120
         TabIndex        =   20
         Top             =   3120
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Room_Charges"
         Height          =   285
         Index           =   7
         Left            =   3120
         TabIndex        =   19
         Top             =   3555
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Hospital_Charges"
         Height          =   285
         Index           =   8
         Left            =   3120
         TabIndex        =   18
         Top             =   4005
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Discount"
         Height          =   285
         Index           =   9
         Left            =   3120
         TabIndex        =   17
         Top             =   4440
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Net_Value"
         Height          =   285
         Index           =   10
         Left            =   3120
         TabIndex        =   16
         Top             =   4875
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Other_Bill_Details"
         Height          =   285
         Index           =   11
         Left            =   3120
         TabIndex        =   15
         Top             =   5325
         Width           =   3375
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Patient Bill ID:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   38
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Patient ID:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   37
         Top             =   915
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Admission ID:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   36
         Top             =   1365
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "DischargeDate:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   35
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Doctor Charges:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   34
         Top             =   2235
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Medicine Charges:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   33
         Top             =   2685
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Services Charges:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   32
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Room Charges:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   31
         Top             =   3555
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Hospital Charges:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   30
         Top             =   4005
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Discount:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   29
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Net Value:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   28
         Top             =   4875
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Notes:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   27
         Top             =   5325
         Width           =   1815
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
      Left            =   840
      TabIndex        =   8
      Top             =   7560
      Width           =   7215
      Begin VB.CommandButton cmdFirst 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   360
         Picture         =   "frmInPatientBill.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   1065
         Picture         =   "frmInPatientBill.frx":5CB8
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   5520
         Picture         =   "frmInPatientBill.frx":6199
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdLast 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   6240
         Picture         =   "frmInPatientBill.frx":6674
         Style           =   1  'Graphical
         TabIndex        =   9
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
         TabIndex        =   13
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
      Height          =   1455
      Left            =   840
      TabIndex        =   0
      Top             =   9000
      Width           =   7215
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   780
         Left            =   1920
         Picture         =   "frmInPatientBill.frx":6B49
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Update"
         Height          =   780
         Left            =   720
         Picture         =   "frmInPatientBill.frx":704D
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   780
         Left            =   5520
         Picture         =   "frmInPatientBill.frx":74FB
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   780
         Left            =   4320
         Picture         =   "frmInPatientBill.frx":7A07
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   780
         Left            =   3120
         Picture         =   "frmInPatientBill.frx":7EAD
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         Height          =   780
         Left            =   1920
         Picture         =   "frmInPatientBill.frx":8366
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         Height          =   780
         Left            =   720
         Picture         =   "frmInPatientBill.frx":880B
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
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
      FormHeightDT    =   11340
      FormWidthDT     =   8850
      FormScaleHeightDT=   10830
      FormScaleWidthDT=   8730
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "IN PATIENT BILL DETAILS"
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
      Index           =   12
      Left            =   1920
      TabIndex        =   39
      Top             =   360
      Width           =   5280
   End
End
Attribute VB_Name = "frmInPatientBill"
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

Private Sub Form_Load()
 
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select Patient_Bill_ID,Patient_ID,Admission_ID,DischargeDate,Doctor_Charges,Medicine_Charges,Services_Charges,Room_Charges,Hospital_Charges,Discount,Net_Value,Other_Bill_Details from Patient_Bill", cnPatients, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
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
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
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
End Sub

