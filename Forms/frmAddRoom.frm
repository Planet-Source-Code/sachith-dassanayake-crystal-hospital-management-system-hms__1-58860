VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmAddRoom 
   BackColor       =   &H00FF8080&
   Caption         =   "Add Room Details"
   ClientHeight    =   6825
   ClientLeft      =   1110
   ClientTop       =   450
   ClientWidth     =   8310
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddRoom.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6825
   ScaleWidth      =   8310
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Search"
      Height          =   780
      Left            =   3720
      Picture         =   "frmAddRoom.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5520
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      DragMode        =   1  'Automatic
      DrawWidth       =   12
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   2775
      Left            =   1080
      Negotiate       =   -1  'True
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      ScaleHeight     =   2715
      ScaleWidth      =   6195
      TabIndex        =   23
      Top             =   2280
      Visible         =   0   'False
      Width           =   6255
      Begin VB.OptionButton optSearch 
         BackColor       =   &H80000018&
         Caption         =   "Room ID"
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   29
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optSearch 
         BackColor       =   &H80000018&
         Caption         =   "Room Type"
         Height          =   375
         Index           =   1
         Left            =   3840
         TabIndex        =   28
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtSearch 
         Height          =   375
         Left            =   960
         TabIndex        =   27
         Text            =   "Enter your Text Here"
         Top             =   1320
         Width           =   4575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000FF&
         Caption         =   "X"
         Height          =   325
         Left            =   5880
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   15
         Width           =   325
      End
      Begin VB.CommandButton cmdFFind 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Search"
         Height          =   780
         Left            =   1800
         Picture         =   "frmAddRoom.frx":5C65
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdFCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   780
         Left            =   3600
         Picture         =   "frmAddRoom.frx":60E8
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000001&
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   80
         Width           =   855
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H80000001&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   6255
      End
   End
   Begin VB.ComboBox cmbRoomType 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   1800
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Record Operation"
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
      Height          =   2295
      Left            =   1440
      TabIndex        =   18
      Top             =   4200
      Width           =   5775
      Begin VB.CommandButton cmdView 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&View All"
         Height          =   780
         Left            =   1080
         Picture         =   "frmAddRoom.frx":65F4
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         Height          =   780
         Left            =   600
         Picture         =   "frmAddRoom.frx":6AAF
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         Height          =   780
         Left            =   1800
         Picture         =   "frmAddRoom.frx":6F5E
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   780
         Left            =   3000
         Picture         =   "frmAddRoom.frx":7403
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   780
         Left            =   4200
         Picture         =   "frmAddRoom.frx":78BC
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   780
         Left            =   3480
         Picture         =   "frmAddRoom.frx":7D62
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Update"
         Height          =   780
         Left            =   600
         Picture         =   "frmAddRoom.frx":826E
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   780
         Left            =   1800
         Picture         =   "frmAddRoom.frx":871C
         Style           =   1  'Graphical
         TabIndex        =   20
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
      Height          =   855
      Left            =   840
      TabIndex        =   17
      Top             =   3120
      Width           =   6855
      Begin VB.CommandButton cmdLast 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   6120
         Picture         =   "frmAddRoom.frx":8C20
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   5400
         Picture         =   "frmAddRoom.frx":90F5
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   945
         Picture         =   "frmAddRoom.frx":95D0
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdFirst 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   240
         Picture         =   "frmAddRoom.frx":9AB1
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1680
         TabIndex        =   6
         Top             =   360
         Width           =   3480
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Room_Description"
      Height          =   285
      Index           =   3
      Left            =   3480
      TabIndex        =   3
      Tag             =   "Chr"
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Room_Type"
      Height          =   285
      Index           =   1
      Left            =   3480
      TabIndex        =   2
      Tag             =   "Chr"
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Room_ID"
      Height          =   285
      Index           =   0
      Left            =   3480
      TabIndex        =   1
      Top             =   1260
      Width           =   3375
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   360
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      FormHeightDT    =   7335
      FormWidthDT     =   8430
      FormScaleHeightDT=   6825
      FormScaleWidthDT=   8310
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   2055
      Left            =   1080
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   6255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "ROOM DETAILS"
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
      Left            =   2760
      TabIndex        =   22
      Top             =   360
      Width           =   3090
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Room Description:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   16
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Room Type:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   15
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Room ID:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "frmAddRoom"
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

Private Sub cmdFCancel_Click()
Picture1.Visible = False
Frame1.Enabled = True
Frame2.Enabled = True
End Sub

Private Sub cmdFFind_Click()
Dim strText As String
Dim SQL As String
'strText = InputBox("Please Enter The patient ID", "Search Patient", "OPID_")

strText = txtSearch
If optSearch(0) = True Then
    SearchFor = "Room_ID"
ElseIf optSearch(1) = True Then
    SearchFor = "Room_Type"
End If

varBookMark = adoPrimaryRS.Bookmark
adoPrimaryRS.MoveFirst

SQL = SearchFor & "=" & "'" & strText & "'"

adoPrimaryRS.Find SQL


If (adoPrimaryRS.BOF = True) Or (adoPrimaryRS.EOF = True) Then
   MsgBox "Record not found", vbInformation, "Search Result"
   adoPrimaryRS.Bookmark = varBookMark
End If
End Sub

Private Sub cmdFind_Click()
Picture1.Visible = True
Frame1.Enabled = False
Frame2.Enabled = False

End Sub

Private Sub cmdView_Click()
frmRoomDetails.Show
End Sub

Private Sub Command1_Click()
Picture1.Visible = False
Frame1.Enabled = True
Frame2.Enabled = True
End Sub

Private Sub Form_Load()
    

    Me.WindowState = vbMaximized



  Call Functions.DisableMenu
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select Room_ID,Room_Type,Room_Description from Room_Details Order by Room_Type", cnPatients, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
    oText.Locked = True
  Next
  Call addRoom
  mbDataChanged = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 22 Then KeyAscii = 0: Exit Sub
    KeyAscii = DataEntryValidation(KeyAscii, ActiveControl.Tag)
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
  lblStatus.Caption = "Room : " & CStr(adoPrimaryRS.AbsolutePosition)
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
  
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    oText.Locked = False
  Next
  txtFields(0).Locked = True
  
    Dim rsRoom As Recordset
    Dim RoomID As String
    
    
    RoomID = Functions.UID(6, "RoomID_")
    
    Set rsRoom = New ADODB.Recordset
    rsRoom.Open "Select * from Bed_Details", cnPatients, adOpenKeyset, adLockPessimistic
              
            While rsRoom.EOF = False
                If rsRoom(0) = RoomID Then
                    RoomID = Functions.UID(6, "RoomID_")
                    rsRoom.MoveFirst
                End If
                rsRoom.MoveNext
            Wend
            
  
            rsRoom.Close
  
  

  
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    txtFields(0) = RoomID
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

 cmbRoomType = txtFields(1)

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
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    oText.Locked = True
  Next
  txtFields(0).Locked = True

  txtFields(1) = cmbRoomType
      
  If txtFields(1) = "" Then
    MsgBox "Please Select the Room type", vbCritical, "Crystal HMS"
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
  cmdFind.Visible = bVal
  cmdView.Visible = bVal
  
  cmbRoomType.Visible = Not bVal
End Sub

Private Sub addRoom()
Dim rsRoomType As Recordset
Set rsRoomType = New ADODB.Recordset

rsRoomType.Open "Select * from Room_Types", cnPatients, adOpenDynamic, adLockReadOnly

While rsRoomType.EOF = False
    cmbRoomType.AddItem rsRoomType(0)
    rsRoomType.MoveNext
Wend
rsRoomType.Close
End Sub
