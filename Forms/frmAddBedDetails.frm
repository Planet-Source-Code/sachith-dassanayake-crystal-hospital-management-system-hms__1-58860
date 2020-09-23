VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmAddBedDetails 
   BackColor       =   &H00FF8080&
   Caption         =   "Bed Details"
   ClientHeight    =   8475
   ClientLeft      =   1110
   ClientTop       =   450
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddBedDetails.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8475
   ScaleWidth      =   8775
   WindowState     =   2  'Maximized
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   7920
      Top             =   6720
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      AutoCenterForm  =   -1  'True
      FormHeightDT    =   8985
      FormWidthDT     =   8895
      FormScaleHeightDT=   8475
      FormScaleWidthDT=   8775
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Bed_Desc"
      Height          =   870
      Index           =   2
      Left            =   2160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Available"
      Height          =   285
      Index           =   3
      Left            =   2160
      TabIndex        =   23
      Top             =   3720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ComboBox cmbWardID 
      Height          =   315
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   3000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.ComboBox cmbRoomID 
      Height          =   315
      Left            =   6360
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   2040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&View"
      Height          =   780
      Left            =   3360
      Picture         =   "frmAddBedDetails.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   780
      Left            =   3360
      Picture         =   "frmAddBedDetails.frx":5C9D
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Update"
      Height          =   780
      Left            =   2160
      Picture         =   "frmAddBedDetails.frx":61A1
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   780
      Left            =   4560
      Picture         =   "frmAddBedDetails.frx":664F
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      Height          =   780
      Left            =   5760
      Picture         =   "frmAddBedDetails.frx":6B5B
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Delete"
      Height          =   780
      Left            =   4560
      Picture         =   "frmAddBedDetails.frx":7001
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Edit"
      Height          =   780
      Left            =   3360
      Picture         =   "frmAddBedDetails.frx":74BA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Add"
      Height          =   780
      Left            =   2160
      Picture         =   "frmAddBedDetails.frx":795F
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton cmdWard 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ward"
      Height          =   825
      Left            =   5160
      Picture         =   "frmAddBedDetails.frx":7E0E
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdRoom 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Room"
      Height          =   780
      Left            =   5160
      Picture         =   "frmAddBedDetails.frx":83B9
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   1320
      Picture         =   "frmAddBedDetails.frx":88AF
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   585
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   2025
      Picture         =   "frmAddBedDetails.frx":8D85
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   585
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   6480
      Picture         =   "frmAddBedDetails.frx":9266
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   585
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   7200
      Picture         =   "frmAddBedDetails.frx":9741
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   585
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Room_Ward_ID"
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   2
      Top             =   2295
      Width           =   2535
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Bed_ID"
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Top             =   1380
      Width           =   2535
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize2 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      FormHeightDT    =   8985
      FormWidthDT     =   8895
      FormScaleHeightDT=   8475
      FormScaleWidthDT=   8775
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Line Line1 
      X1              =   -600
      X2              =   9240
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   3735
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   8415
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   2055
      Left            =   1800
      Top             =   6120
      Width           =   5415
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FFFFFF&
      Height          =   855
      Left            =   840
      Top             =   5040
      Width           =   7335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "BED DETAILS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2760
      TabIndex        =   24
      Top             =   240
      Width           =   2865
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Left            =   2760
      TabIndex        =   6
      Top             =   5280
      Width           =   3480
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Bed Description:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   19
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Room / Ward ID:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   18
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Bed ID:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
   End
End
Attribute VB_Name = "frmAddBedDetails"
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
Dim ChkClick As Integer

Private Sub cmbRoomID_Click()
txtFields(1) = cmbRoomID
ChkClick = 1
End Sub

Private Sub cmbWardID_Click()
txtFields(1) = cmbWardID
ChkClick = 2
End Sub

Private Sub cmdRoom_Click()
frmRoomDetails.Show
End Sub

Private Sub cmdView_Click()
frmBedDetails.Show
End Sub

Private Sub cmdWard_Click()
frmWardDetails.Show
End Sub

Private Sub Form_Activate()
Call Functions.DisableMenu
End Sub

Private Sub Form_Load()

  Call Functions.DisableMenu
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select Bed_ID,Room_Ward_ID,Bed_Desc,Available from Bed_Details Order by Bed_ID", cnPatients, adOpenStatic, adLockOptimistic
  Dim i
  i = 0
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
    oText.Locked = True
    oText.TabIndex = i
    i = i + 1
  Next
    Call FillData
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
  lblStatus.Caption = "Bed: " & CStr(adoPrimaryRS.AbsolutePosition)
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
  
    Dim rsBed As Recordset
    Dim BedID As String
    Dim oText As TextBox
    For Each oText In Me.txtFields
        oText.Locked = False
    Next
    txtFields(0).Locked = True
    txtFields(1).Locked = True
    
    BedID = Functions.UID(6, "BedID_")
    
    Set rsBed = New ADODB.Recordset
    rsBed.Open "Select * from Bed_Details", cnPatients, adOpenKeyset, adLockPessimistic
              
            While rsBed.EOF = False
                If rsBed(0) = BedID Then
                    BedID = Functions.UID(6, "BedID_")
                    rsBed.MoveFirst
                End If
                rsBed.MoveNext
            Wend
            
  
            rsBed.Close
    
    
        
  
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    txtFields(0) = BedID
    lblStatus.Caption = "Add record"
    mbAddNewFlag = True
    SetButtons False
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description, vbCritical, "An Error Occured"
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
  MsgBox Err.Description, vbCritical, "An Error Occured"
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
    For Each oText In Me.txtFields
        oText.Locked = False
    Next
    txtFields(0).Locked = True
    txtFields(1).Locked = True

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description, vbCritical, "An Error Occured"
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next
    Dim oText As TextBox
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
    Dim rsCheckBed As Recordset
    Set rsCheckBed = New ADODB.Recordset
    txtFields(3) = True
    
    If ChkClick = 1 Then
    rsCheckBed.Open "Select * from Bed_Details where Room_Ward_ID  = '" & txtFields(1) & "'", cnPatients, adOpenDynamic, adLockReadOnly
    If rsCheckBed.EOF = False Then
        MsgBox "This Room is alredy assigned with a bed", vbCritical, "Crystal HMS"
        Exit Sub
    End If
    rsCheckBed.Close
    End If
    
    Dim oText As TextBox
    For Each oText In Me.txtFields
        oText.Locked = True
    Next


    If Trim(txtFields(1)) = "" Then
    MsgBox "Please enter Room ID or Ward ID", vbCritical, "Crystal HMS"
    txtFields(1).SetFocus
    SendKeys ("Shift" + "Home")
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
  MsgBox Err.Description, vbCritical, "An Error Occured"
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
  MsgBox Err.Description, vbCritical, "An Error Occured"
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  adoPrimaryRS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description, vbCritical, "An Error Occured"
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
  MsgBox Err.Description, vbCritical, "An Error Occured"
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
  MsgBox Err.Description, vbCritical, "An Error Occured"
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
  cmdRoom.Visible = Not bVal
  cmdWard.Visible = Not bVal
  cmdView.Visible = bVal
  
  cmbRoomID.Visible = Not bVal
  cmbWardID.Visible = Not bVal
    
End Sub

Private Sub FillData()
Dim rsRoom As Recordset
Dim rsWard As Recordset

Set rsRoom = New ADODB.Recordset
Set rsWard = New ADODB.Recordset


rsRoom.Open "Select * from Room_Details", cnPatients, adOpenDynamic, adLockReadOnly
rsWard.Open "Select * from Ward_Details", cnPatients, adOpenDynamic, adLockReadOnly

While rsRoom.EOF = False
cmbRoomID.AddItem rsRoom(0)

rsRoom.MoveNext
Wend


While rsWard.EOF = False
cmbWardID.AddItem rsWard(0)
rsWard.MoveNext
Wend

rsRoom.Close
rsWard.Close

End Sub

