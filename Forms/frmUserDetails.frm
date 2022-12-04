VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmUserDetails 
   BackColor       =   &H00FF8080&
   Caption         =   "User Details"
   ClientHeight    =   9765
   ClientLeft      =   4500
   ClientTop       =   1230
   ClientWidth     =   8640
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserDetails.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9765
   ScaleWidth      =   8640
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdUsertype 
      Caption         =   "---"
      Height          =   255
      Left            =   5280
      TabIndex        =   40
      Top             =   6000
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF8080&
      Caption         =   "Controls"
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
      Height          =   7455
      Left            =   6480
      TabIndex        =   31
      Top             =   840
      Width           =   1935
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         Height          =   780
         Left            =   360
         Picture         =   "frmUserDetails.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         Height          =   780
         Left            =   360
         Picture         =   "frmUserDetails.frx":12F5
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   780
         Left            =   360
         Picture         =   "frmUserDetails.frx":17CC
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   780
         Left            =   360
         Picture         =   "frmUserDetails.frx":1C8F
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   780
         Left            =   360
         Picture         =   "frmUserDetails.frx":2135
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   6240
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Update"
         Height          =   780
         Left            =   360
         Picture         =   "frmUserDetails.frx":2641
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   780
         Left            =   360
         Picture         =   "frmUserDetails.frx":2AEF
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1560
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdViewAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&View All"
         Height          =   780
         Left            =   360
         Picture         =   "frmUserDetails.frx":2FF3
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   5040
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Navigation"
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
      Left            =   240
      TabIndex        =   25
      Top             =   8520
      Width           =   8175
      Begin VB.CommandButton cmdLast 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   6840
         Picture         =   "frmUserDetails.frx":34AE
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   6120
         Picture         =   "frmUserDetails.frx":3983
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   1665
         Picture         =   "frmUserDetails.frx":3E5E
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdFirst 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   960
         Picture         =   "frmUserDetails.frx":433F
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   2520
         TabIndex        =   30
         Top             =   360
         Width           =   3360
      End
   End
   Begin VB.TextBox txtUserPass 
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
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "•"
      TabIndex        =   22
      Top             =   7680
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Notes"
      Height          =   285
      Index           =   10
      Left            =   2400
      TabIndex        =   21
      Top             =   5160
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Status"
      Height          =   285
      Index           =   9
      Left            =   2400
      TabIndex        =   19
      Top             =   4620
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Telephone"
      Height          =   285
      Index           =   8
      Left            =   2400
      TabIndex        =   17
      Top             =   4080
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Email"
      Height          =   285
      Index           =   7
      Left            =   2400
      TabIndex        =   15
      Top             =   3480
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Address"
      Height          =   285
      Index           =   6
      Left            =   2400
      TabIndex        =   13
      Top             =   2940
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Gender"
      Height          =   285
      Index           =   5
      Left            =   2400
      TabIndex        =   11
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Last_Name"
      Height          =   285
      Index           =   4
      Left            =   2400
      TabIndex        =   9
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "First_Name"
      Height          =   285
      Index           =   3
      Left            =   2400
      TabIndex        =   7
      Top             =   1260
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "User_Type"
      Height          =   285
      Index           =   2
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   6000
      Width           =   2775
   End
   Begin VB.TextBox txtFields 
      DataField       =   "User_Password"
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
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2400
      PasswordChar    =   "•"
      TabIndex        =   3
      Top             =   7080
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "User_Name"
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   2400
      TabIndex        =   1
      Top             =   6540
      Width           =   3375
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   240
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      FormHeightDT    =   10275
      FormWidthDT     =   8760
      FormScaleHeightDT=   9765
      FormScaleWidthDT=   8640
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "USER DETAILS"
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
      Left            =   3120
      TabIndex        =   24
      Top             =   240
      Width           =   2925
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   4695
      Left            =   240
      Top             =   960
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   2535
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   5760
      Width           =   6015
   End
   Begin VB.Label lblUserPass 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Re type Password"
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
      TabIndex        =   23
      Top             =   7680
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Notes:"
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
      Index           =   10
      Left            =   840
      TabIndex        =   20
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Status:"
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
      Index           =   9
      Left            =   840
      TabIndex        =   18
      Top             =   4680
      Width           =   675
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Telephone:"
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
      Index           =   8
      Left            =   840
      TabIndex        =   16
      Top             =   4080
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Email:"
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
      Index           =   7
      Left            =   840
      TabIndex        =   14
      Top             =   3480
      Width           =   600
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Address:"
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
      Index           =   6
      Left            =   840
      TabIndex        =   12
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Gender:"
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
      Index           =   5
      Left            =   840
      TabIndex        =   10
      Top             =   2400
      Width           =   765
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Last Name:"
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
      Index           =   4
      Left            =   840
      TabIndex        =   8
      Top             =   1800
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "First Name:"
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
      Index           =   3
      Left            =   840
      TabIndex        =   6
      Top             =   1200
      Width           =   1125
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "User Type:"
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
      Index           =   2
      Left            =   480
      TabIndex        =   4
      Top             =   6000
      Width           =   1050
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "User Password:"
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
      Index           =   1
      Left            =   480
      TabIndex        =   2
      Top             =   7080
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "User Name:"
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
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   6600
      Width           =   1125
   End
End
Attribute VB_Name = "frmUserDetails"
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

Private Sub cmdUsertype_Click()
frmUserTypes.Show
End Sub

Private Sub Form_Load()
  Me.WindowState = vbMaximized
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select User_Name,User_Password,User_Type,First_Name,Last_Name,Gender,Address,Email,Telephone,Status,Notes from User_Details", cnPatients, adOpenDynamic, adLockPessimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
    oText.Enabled = False
  Next

  mbDataChanged = False
  Call Functions.DisableMenu
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
    oText.Enabled = True
  Next
  txtFields(2).Locked = True
  
  
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
     Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
   
    oText.Enabled = True
  Next
  
  
  
  
    txtFields(1) = Functions.Decrypt(txtFields(1))
  txtUserPass = txtFields(1)
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
   
    oText.Enabled = False
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
  Dim oText As TextBox

    If txtFields(2) = "" Then
        MsgBox "Please Select the User type", vbCritical, "Error"
        cmdUsertype.SetFocus
        Exit Sub
    End If
    If txtFields(0) = "" Then
        MsgBox "Please Enter User Name", vbCritical, "Error"
        txtFields(0).SetFocus
        Exit Sub
    End If
        


    'txtFields(1) = Functions.Encrypt(txtFields(1))
    If txtFields(1) <> txtUserPass Then
        MsgBox "Password does not match", vbCritical
        txtUserPass.SetFocus
        SendKeys "{Home}+{End}"
        Exit Sub
    End If
    
    txtFields(1) = Functions.Encrypt(txtFields(1))

  adoPrimaryRS.UpdateBatch adAffectAll
  
     
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    oText.Enabled = False
  Next
  
    txtUserPass = ""
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
  
  txtUserPass.Visible = Not bVal
  txtFields(1).Visible = Not bVal
  lblLabels(1).Visible = Not bVal
  lblUserPass.Visible = Not bVal
  
  cmdViewAll.Visible = bVal
  cmdUsertype.Visible = Not bVal
  
  
End Sub


