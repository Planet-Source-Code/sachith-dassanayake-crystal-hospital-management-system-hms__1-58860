VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmDoctorSchedule 
   BackColor       =   &H00FF8080&
   Caption         =   "Doctors Channeling Schedule Details"
   ClientHeight    =   10860
   ClientLeft      =   1110
   ClientTop       =   450
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDoctorSchedule.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10860
   ScaleWidth      =   10020
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
      Height          =   1215
      Left            =   1440
      TabIndex        =   34
      Top             =   8040
      Width           =   7215
      Begin VB.CommandButton cmdFirst 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   360
         Picture         =   "frmDoctorSchedule.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1065
         Picture         =   "frmDoctorSchedule.frx":5CB8
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   5520
         Picture         =   "frmDoctorSchedule.frx":6199
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdLast 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   6240
         Picture         =   "frmDoctorSchedule.frx":6674
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1800
         TabIndex        =   39
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
      Height          =   1215
      Left            =   1440
      TabIndex        =   26
      Top             =   9480
      Width           =   7215
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   780
         Left            =   5760
         Picture         =   "frmDoctorSchedule.frx":6B49
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   780
         Left            =   4440
         Picture         =   "frmDoctorSchedule.frx":7055
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   780
         Left            =   3120
         Picture         =   "frmDoctorSchedule.frx":74FB
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         Height          =   780
         Left            =   1800
         Picture         =   "frmDoctorSchedule.frx":79B4
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   1800
         Picture         =   "frmDoctorSchedule.frx":7E59
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Update"
         Height          =   780
         Left            =   480
         Picture         =   "frmDoctorSchedule.frx":835D
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   480
         Picture         =   "frmDoctorSchedule.frx":880B
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "Doctor Details"
      ForeColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   360
      TabIndex        =   10
      Top             =   1560
      Width           =   6375
      Begin VB.TextBox txtDay 
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   2760
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Schedule_ID"
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   18
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Doctor_AvaiDate"
         Height          =   285
         Index           =   4
         Left            =   2160
         TabIndex        =   14
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Schedule_Notes"
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   13
         Top             =   3360
         Width           =   3375
      End
      Begin VB.ComboBox cmbDoctorID 
         Height          =   315
         ItemData        =   "frmDoctorSchedule.frx":8CBA
         Left            =   2160
         List            =   "frmDoctorSchedule.frx":8CBC
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   960
         Width           =   3375
      End
      Begin VB.CommandButton cmdShowAll 
         Caption         =   "..."
         Height          =   300
         Left            =   5640
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtTempDocID 
         Height          =   285
         Left            =   2160
         TabIndex        =   19
         Top             =   960
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Doctor_ID"
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   17
         Top             =   960
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker DTPin 
         Height          =   375
         Left            =   2160
         TabIndex        =   41
         Top             =   1560
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   20709379
         UpDown          =   -1  'True
         CurrentDate     =   38380
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Doctor_In"
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   16
         Top             =   1560
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker DTPout 
         Height          =   375
         Left            =   2160
         TabIndex        =   42
         Top             =   2160
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   20709379
         UpDown          =   -1  'True
         CurrentDate     =   38380
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Doctor_Out"
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   15
         Top             =   2160
         Width           =   3375
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Schedule_ID:"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Doctor ID:"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Time In:"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   23
         Top             =   1545
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Time Out:"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   22
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Available Days:"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   21
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Schedule Notes:"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   20
         Top             =   3360
         Width           =   1815
      End
   End
   Begin VB.Frame frameDays 
      BackColor       =   &H00FF8080&
      Caption         =   "Available Days"
      ForeColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   7440
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
      Begin VB.CheckBox chkDay 
         BackColor       =   &H00FF8080&
         Caption         =   "Wednesday"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   3
         Left            =   360
         TabIndex        =   8
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CheckBox chkDay 
         BackColor       =   &H00FF8080&
         Caption         =   "Saturday"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   6
         Left            =   360
         TabIndex        =   7
         Top             =   3240
         Width           =   1215
      End
      Begin VB.CheckBox chkDay 
         BackColor       =   &H00FF8080&
         Caption         =   "Friday"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   5
         Left            =   360
         TabIndex        =   6
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CheckBox chkDay 
         BackColor       =   &H00FF8080&
         Caption         =   "Thursday"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Top             =   2280
         Width           =   1215
      End
      Begin VB.CheckBox chkDay 
         BackColor       =   &H00FF8080&
         Caption         =   "Tuesday"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   2
         Left            =   360
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CheckBox chkDay 
         BackColor       =   &H00FF8080&
         Caption         =   "Monday"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin VB.CheckBox chkDay 
         BackColor       =   &H00FF8080&
         Caption         =   "Sunday"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Height          =   1305
      Left            =   480
      TabIndex        =   0
      Top             =   6360
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   2302
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
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
      FormHeightDT    =   11370
      FormWidthDT     =   10140
      FormScaleHeightDT=   10860
      FormScaleWidthDT=   10020
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "DOCTOR APPOINTMENT SCHEDULING"
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
      Index           =   6
      Left            =   1080
      TabIndex        =   9
      Top             =   480
      Width           =   7545
   End
End
Attribute VB_Name = "frmDoctorSchedule"
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
Dim strDays As String

Private Sub chkDay_Click(Index As Integer)

Select Case (Index)

Case 0
If chkDay(0).Value = 1 Then
    strDays = strDays & "Sun,"
    'txtFields(4) = strDays
    txtDay = strDays
ElseIf chkDay(0).Value = 0 Then
 strDays = Replace(strDays, "Sun,", "")
 'txtFields(4).Text = strDays
 txtDay = strDays
End If

Case 1
If chkDay(1).Value = 1 Then
    strDays = strDays & "Mon,"
    'txtFields(4) = strDays
    txtDay = strDays
ElseIf chkDay(1).Value = 0 Then
 strDays = Replace(strDays, "Mon,", "")
 'txtFields(4).Text = strDays
 txtDay = strDays
End If


Case 2
If chkDay(2).Value = 1 Then
    strDays = strDays & "Tue,"
    'txtFields(4) = strDays
    txtDay = strDays
ElseIf chkDay(2).Value = 0 Then
 strDays = Replace(strDays, "Tue,", "")
 'txtFields(4).Text = strDays
 txtDay = strDays
End If


Case 3
If chkDay(3).Value = 1 Then
    strDays = strDays & "Wed,"
    'txtFields(4) = strDays
    txtDay = strDays
ElseIf chkDay(3).Value = 0 Then
 strDays = Replace(strDays, "Wed,", "")
 'txtFields(4).Text = strDays
 txtDay = strDays
End If


Case 4
If chkDay(4).Value = 1 Then
    strDays = strDays & "Thu,"
    'txtFields(4) = strDays
    txtDay = strDays
ElseIf chkDay(4).Value = 0 Then
 strDays = Replace(strDays, "Thu,", "")
 'txtFields(4).Text = strDays
 txtDay = strDays
End If


Case 5
If chkDay(5).Value = 1 Then
    strDays = strDays & "Fri,"
    'txtFields(4) = strDays
    txtDay = strDays
ElseIf chkDay(5).Value = 0 Then
 strDays = Replace(strDays, "Fri,", "")
 'txtFields(4).Text = strDays
 txtDay = strDays
End If


Case 6
If chkDay(6).Value = 1 Then
    strDays = strDays & "Sat,"
    'txtFields(4) = strDays
    txtDay = strDays
ElseIf chkDay(6).Value = 0 Then
 strDays = Replace(strDays, "Sat,", "")
 'txtFields(4).Text = strDays
 txtDay = strDays
End If


Case Else

End Select


End Sub





Private Sub cmbDoctorID_Click()
txtTempDocID = cmbDoctorID.Text
End Sub


Private Sub cmdShowAll_Click()
frmDoctorDetails.Show
End Sub



Private Sub Form_Activate()
'cmbDoctorID = PvtDocID
End Sub

Private Sub Form_Load()
  Call Functions.DisableMenu
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "SHAPE {select Schedule_ID,Doctor_ID,Doctor_In,Doctor_Out,Doctor_AvaiDate,Schedule_Notes from Doctor_Schedule_Details} AS ParentCMD APPEND ({select Schedule_ID,Doctor_ID,Doctor_In,Doctor_Out,Doctor_AvaiDate,Schedule_Notes from Doctor_Schedule_Details Order by Schedule_ID } AS ChildCMD RELATE Doctor_ID TO Doctor_ID) AS ChildCMD", cnPatients, adOpenDynamic, adLockPessimistic

  Dim oText As TextBox
  Dim oChk As CheckBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
    oText.Locked = True
  Next

  cmbDoctorID.Visible = False
  frameDays.Visible = False
  
  Set grdDataGrid.DataSource = adoPrimaryRS("ChildCMD").UnderlyingValue

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
  
  Dim SchedID As String
  Dim flag As Boolean
  Dim rsSched As Recordset
  Dim i As Integer
  Set rsSched = New ADODB.Recordset
  
  
   For i = 0 To 6
        chkDay(i).Value = 0
    Next
   
        SchedID = Functions.UID(6, "SchedID_")
        rsSched.Open "Select * from Doctor_Schedule_Details", cnPatients, adOpenKeyset, adLockPessimistic
              
            While rsSched.EOF = False
                If rsSched(0) = SchedID Then
                   SchedID = Functions.UID(6, "SchedID_")
                   rsSched.MoveFirst
                    flag = True
                Else
                    flag = False
                End If
                rsSched.MoveNext
            Wend
            
  
            rsSched.Close

    cmbDoctorID.clear
    rsSched.Open "SELECT * FROM Doctor_Details", cnPatients, adOpenKeyset, adLockPessimistic
    i = 0
    ' Add ID's to Combo Box
    
    While rsSched.EOF = False
        cmbDoctorID.AddItem rsSched(0)
        cmbDoctorID.ListIndex = i
        i = i + 1
        rsSched.MoveNext
    Wend
   cmbDoctorID.Text = cmbDoctorID.List(0)
   rsSched.Close
            
  
   
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    
    cmbDoctorID.Visible = True
    txtFields(1).Visible = False
    txtFields(1) = cmbDoctorID.Text
    txtTempDocID = cmbDoctorID.Text
    
    lblStatus.Caption = "Add record"
    frameDays.Visible = True
    txtFields(0).Text = SchedID
    
    Dim oText As TextBox
    
    For Each oText In Me.txtFields
        oText.Locked = False
    Next
    
    txtFields(0).Locked = True
    txtFields(1).Locked = True
    
    
    
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
  Set grdDataGrid.DataSource = Nothing
  adoPrimaryRS.Requery
  Set grdDataGrid.DataSource = adoPrimaryRS("ChildCMD").UnderlyingValue
  
  
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  Set grdDataGrid.DataSource = Nothing
  adoPrimaryRS.Requery
  Set grdDataGrid.DataSource = adoPrimaryRS("ChildCMD").UnderlyingValue
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  Dim rsSched As Recordset
  Dim i As Integer
  Set rsSched = New ADODB.Recordset

   Dim oText As TextBox
   Dim strDays As String
   Dim arrDays() As String
  On Error GoTo EditErr
    cmbDoctorID.clear
  
    rsSched.Open "SELECT * FROM Doctor_Details", cnPatients, adOpenKeyset, adLockPessimistic
    i = 0
    ' Add ID's to Combo Box
    
    While rsSched.EOF = False
        cmbDoctorID.AddItem rsSched(0)
        cmbDoctorID.ListIndex = i
        i = i + 1
        rsSched.MoveNext
    Wend
    rsSched.Close
    
    strDays = txtFields(4)
    arrDays() = Split(strDays, ",")
    
    For i = 0 To 6
        chkDay(i).Value = 0
    Next
    
    For i = 0 To UBound(arrDays)
               
        If arrDays(i) = "Sun" Then
            chkDay(0).Value = 1
        End If
        If arrDays(i) = "Mon" Then
            chkDay(1).Value = 1
        End If
        If arrDays(i) = "Tue" Then
           chkDay(2).Value = 1
        End If
        If arrDays(i) = "Wed" Then
            chkDay(3).Value = 1
        End If
        If arrDays(i) = "Thu" Then
            chkDay(4).Value = 1
        End If
        If arrDays(i) = "Fri" Then
            chkDay(5).Value = 1
        End If
        If arrDays(i) = "Sat" Then
            chkDay(6).Value = 1
        End If
        
    Next i
    
    
    

  lblStatus.Caption = "Edit record"
    
    frameDays.Visible = True
    For Each oText In Me.txtFields
        oText.Locked = False
    Next
    txtFields(0).Locked = True
  
  
  cmbDoctorID.Visible = True
  cmbDoctorID.Text = txtFields(1).Text
  txtTempDocID = cmbDoctorID.Text
  DTPin.Value = txtFields(2)
  DTPout.Value = txtFields(3)
  
  
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  
  On Error GoTo CancelError
  
    frameDays.Visible = False
    If Right(txtFields(4), 1) = "," Then
        txtFields(4).Text = Left(txtFields(4), Len(txtFields(4)) - 1)
    End If
    
    Dim oText As TextBox
    For Each oText In Me.txtFields
        oText.Locked = True
    Next
    txtFields(0).Locked = True
    txtFields(1).Visible = True
    cmbDoctorID.Visible = False
  
  

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
    
 
Exit Sub

CancelError:
  MsgBox Err.Description

End Sub

Private Sub cmdUpdate_Click()
Dim oText As TextBox

  On Error GoTo UpdateErr
  
  txtFields(1) = cmbDoctorID.Text
  txtFields(4) = txtDay
  txtFields(2) = Format(DTPin.Value, "Short time")
  txtFields(3) = Format(DTPout.Value, "short time")
  
  MsgBox txtFields(2)

    If Right(txtFields(4), 1) = "," Then
        txtFields(4).Text = Left(txtFields(4), Len(txtFields(4)) - 1)
    End If
    
    If Trim(txtFields(2)) = "" Or Trim(txtFields(3)) = "" Or txtFields(4) = "" Then
        MsgBox "Please fill all the relavent fields", vbCritical, "An Error Occured"
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
  
  For Each oText In Me.txtFields
    oText.Locked = True
  Next
  txtFields(1) = txtTempDocID
  txtFields(1).Visible = True
  cmbDoctorID.Visible = False
  
  
  frameDays.Visible = False
  
  
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
  cmdShowAll.Visible = Not bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
  txtFields(4).Locked = Not bVal
  txtFields(1).Locked = Not bVal
  txtTempDocID.Visible = Not bVal
  
  txtDay.Visible = Not bVal
  
  DTPin.Visible = Not bVal
  DTPout.Visible = Not bVal
End Sub


