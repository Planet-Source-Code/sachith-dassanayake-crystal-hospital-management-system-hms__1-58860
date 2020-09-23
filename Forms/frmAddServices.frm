VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmAddServices 
   BackColor       =   &H00FF8080&
   Caption         =   "Services"
   ClientHeight    =   8565
   ClientLeft      =   4140
   ClientTop       =   1695
   ClientWidth     =   7875
   Icon            =   "frmAddServices.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   7875
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   6600
      Picture         =   "frmAddServices.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   585
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   5880
      Picture         =   "frmAddServices.frx":5CB7
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   585
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   1425
      Picture         =   "frmAddServices.frx":6192
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   585
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00FFFFFF&
      Height          =   540
      Left            =   720
      Picture         =   "frmAddServices.frx":6673
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   585
   End
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
      Left            =   1800
      Picture         =   "frmAddServices.frx":6B49
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Edit"
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
      Left            =   3000
      Picture         =   "frmAddServices.frx":6FF8
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFFFF&
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
      Height          =   780
      Left            =   4200
      Picture         =   "frmAddServices.frx":749D
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6360
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
      Left            =   5400
      Picture         =   "frmAddServices.frx":7956
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6360
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
      Left            =   4200
      Picture         =   "frmAddServices.frx":7DFC
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7320
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Update"
      Height          =   780
      Left            =   1800
      Picture         =   "frmAddServices.frx":8308
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Cancel"
      Height          =   780
      Left            =   3000
      Picture         =   "frmAddServices.frx":87B6
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdView 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&View"
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
      Left            =   3000
      Picture         =   "frmAddServices.frx":8CBA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Service_Notes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Index           =   4
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   3555
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Duration_Of_Service"
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
      Index           =   3
      Left            =   3120
      TabIndex        =   7
      Tag             =   "Num"
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Charge_For_Service"
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
      Index           =   2
      Left            =   3120
      TabIndex        =   5
      Tag             =   "amt"
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Channel_Service"
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
      Left            =   3120
      TabIndex        =   3
      Tag             =   "Chr"
      Top             =   1875
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Channel_Service_ID"
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
      Index           =   0
      Left            =   3120
      TabIndex        =   1
      Top             =   1320
      Width           =   3375
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
      FormHeightDT    =   9075
      FormWidthDT     =   7995
      FormScaleHeightDT=   8565
      FormScaleWidthDT=   7875
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "MEDICAL SERVICES"
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
      Index           =   5
      Left            =   2160
      TabIndex        =   23
      Top             =   360
      Width           =   3960
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   3855
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   6255
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   975
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   5040
      Width           =   7335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   2175
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   6120
      Width           =   6855
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2250
      TabIndex        =   10
      Top             =   5280
      Width           =   3360
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Additional Notes:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   1200
      TabIndex        =   8
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Duration (min):"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   6
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Amount / Rate:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   4
      Top             =   2445
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Service Name:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Service ID:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "frmAddServices"
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

Private Sub cmdView_Click()
frmService.Show
End Sub

Private Sub Form_Load()
  
    Me.WindowState = vbMaximized
  
  
  
  
  Call Functions.DisableMenu
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select Channel_Service_ID,Channel_Service,Charge_For_Service,Duration_Of_Service,Service_Notes from Services Order by Channel_Service_ID", cnPatients, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next

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
  Dim SerID As String
  Dim rsSer As Recordset
  Set rsSer = New ADODB.Recordset
  On Error GoTo AddErr
  
            SerID = Functions.UID(6, "SerID_")
            rsSer.Open "Select * from Services", cnPatients, adOpenKeyset, adLockPessimistic
              
            While rsSer.EOF = False
                If rsSer(0) = SerID Then
                    SerID = Functions.UID(6, "SerID_")
                    flag = True
                Else
                    flag = False
                End If
                rsSer.MoveNext
            Wend
            
  
            rsSer.Close
  
  
  
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    txtFields(0).Text = SerID
    txtFields(0).Enabled = False
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
  txtFields(0).Enabled = False
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
  cmdView.Visible = bVal
End Sub

