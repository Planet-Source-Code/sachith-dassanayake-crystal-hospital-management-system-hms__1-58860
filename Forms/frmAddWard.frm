VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmAddWard 
   BackColor       =   &H00FF8080&
   Caption         =   "Add Ward Details"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7920
   Icon            =   "frmAddWard.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   7920
   WindowState     =   2  'Maximized
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
      Left            =   600
      TabIndex        =   17
      Top             =   3840
      Width           =   6855
      Begin VB.CommandButton cmdFirst 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   240
         Picture         =   "frmAddWard.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   945
         Picture         =   "frmAddWard.frx":5CB8
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   5400
         Picture         =   "frmAddWard.frx":6199
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdLast 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   6120
         Picture         =   "frmAddWard.frx":6674
         Style           =   1  'Graphical
         TabIndex        =   18
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
         TabIndex        =   22
         Top             =   360
         Width           =   3480
      End
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
      Left            =   1200
      TabIndex        =   8
      Top             =   4920
      Width           =   5775
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
         Left            =   1800
         Picture         =   "frmAddWard.frx":6B49
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Update"
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
         Left            =   600
         Picture         =   "frmAddWard.frx":704D
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
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
         Left            =   3000
         Picture         =   "frmAddWard.frx":74FB
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1320
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
         Left            =   4200
         Picture         =   "frmAddWard.frx":7A07
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
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
         Left            =   3000
         Picture         =   "frmAddWard.frx":7EAD
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         Height          =   780
         Left            =   1800
         Picture         =   "frmAddWard.frx":8366
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         Height          =   780
         Left            =   600
         Picture         =   "frmAddWard.frx":880B
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdView 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&View All"
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
         Picture         =   "frmAddWard.frx":8CBA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ward_Desc"
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
      Left            =   3360
      TabIndex        =   7
      Tag             =   "Chr"
      Top             =   3120
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ward_Rate"
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
      Left            =   3360
      TabIndex        =   5
      Tag             =   "Amt"
      Top             =   2565
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ward_Name"
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
      Left            =   3360
      TabIndex        =   3
      Tag             =   "Chr"
      Top             =   1875
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Ward_ID"
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
      Left            =   3360
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
      FormHeightDT    =   8115
      FormWidthDT     =   8040
      FormScaleHeightDT=   7605
      FormScaleWidthDT=   7920
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   2775
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   6255
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "WARD DETAILS"
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
      Left            =   2520
      TabIndex        =   23
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Ward Description:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1440
      TabIndex        =   6
      Top             =   3120
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Ward Rate:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   4
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Ward Type:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Ward ID:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   1320
      Width           =   1815
   End
End
Attribute VB_Name = "frmAddWard"
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
frmWardDetails.Show
End Sub

Private Sub cmdViewAll_Click()
frmWardDetails.Show
End Sub

Private Sub Form_Load()
    

    Me.WindowState = vbMaximized



  Call Functions.DisableMenu
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select Ward_ID,Ward_Name,Ward_Rate,Ward_Desc from Ward_Details Order by Ward_Name", cnPatients, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
    oText.Locked = True
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
  lblStatus.Caption = "Ward : " & CStr(adoPrimaryRS.AbsolutePosition)
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
  
    Dim rsWard As Recordset
    Dim WardID As String
    
    
    WardID = Functions.UID(6, "WardID_")
    
    Set rsWard = New ADODB.Recordset
    rsWard.Open "Select * from Bed_Details", cnPatients, adOpenKeyset, adLockPessimistic
              
            While rsWard.EOF = False
                If rsWard(0) = WardID Then
                    WardID = Functions.UID(6, "WardID_")
                    rsWard.MoveFirst
                End If
                rsWard.MoveNext
            Wend
            
  
            rsWard.Close
  
  

  
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    txtFields(0) = WardID
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
End Sub


