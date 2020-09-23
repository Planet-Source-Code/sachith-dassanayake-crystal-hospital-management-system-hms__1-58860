VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmAddDoctorDetails 
   BackColor       =   &H00FF8080&
   Caption         =   "Doctor Details"
   ClientHeight    =   10200
   ClientLeft      =   3510
   ClientTop       =   450
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddDoctorDetails.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10200
   ScaleWidth      =   9150
   WindowState     =   2  'Maximized
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   8520
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      FormHeightDT    =   10710
      FormWidthDT     =   9270
      FormScaleHeightDT=   10200
      FormScaleWidthDT=   9150
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF8080&
      Caption         =   "Controls"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2295
      Left            =   120
      TabIndex        =   44
      Top             =   7800
      Width           =   8895
      Begin VB.CommandButton cmdViewAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&View All"
         Height          =   780
         Left            =   3480
         Picture         =   "frmAddDoctorDetails.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         Height          =   780
         Left            =   2160
         Picture         =   "frmAddDoctorDetails.frx":12FD
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         Height          =   780
         Left            =   3480
         Picture         =   "frmAddDoctorDetails.frx":17B0
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   780
         Left            =   4800
         Picture         =   "frmAddDoctorDetails.frx":1C87
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   780
         Left            =   6120
         Picture         =   "frmAddDoctorDetails.frx":214A
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   780
         Left            =   4800
         Picture         =   "frmAddDoctorDetails.frx":25F0
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Update"
         Height          =   780
         Left            =   2160
         Picture         =   "frmAddDoctorDetails.frx":2AFC
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   780
         Left            =   3480
         Picture         =   "frmAddDoctorDetails.frx":2FAA
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Navigation"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   43
      Top             =   6840
      Width           =   8895
      Begin VB.CommandButton cmdLast 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   7200
         Picture         =   "frmAddDoctorDetails.frx":34AE
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   6480
         Picture         =   "frmAddDoctorDetails.frx":3983
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   2025
         Picture         =   "frmAddDoctorDetails.frx":3E5E
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdFirst 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   1320
         Picture         =   "frmAddDoctorDetails.frx":433F
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Height          =   525
         Left            =   2880
         TabIndex        =   19
         Top             =   240
         Width           =   3360
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "Personal Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   120
      TabIndex        =   28
      Top             =   1200
      Width           =   8895
      Begin VB.ComboBox cmbGender 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmAddDoctorDetails.frx":4815
         Left            =   2040
         List            =   "frmAddDoctorDetails.frx":481F
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Doctor_Specialization"
         Height          =   285
         Index           =   9
         Left            =   6360
         TabIndex        =   10
         Tag             =   "Chr"
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Doctor_Qualfication"
         Height          =   285
         Index           =   8
         Left            =   6360
         TabIndex        =   9
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Doctor_Address"
         Height          =   1005
         Index           =   5
         Left            =   2040
         TabIndex        =   8
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Doctor_HPhone"
         Height          =   285
         Index           =   6
         Left            =   6360
         TabIndex        =   5
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Doctor_MPhone"
         Height          =   285
         Index           =   7
         Left            =   6360
         TabIndex        =   7
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Doctor_FName"
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   2
         Tag             =   "Chr"
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Doctor_LName"
         Height          =   285
         Index           =   2
         Left            =   6360
         TabIndex        =   3
         Tag             =   "Chr"
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Doctor_Sex"
         Height          =   285
         Index           =   3
         Left            =   2040
         TabIndex        =   4
         Tag             =   "Chr"
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Doctor_NID"
         Height          =   285
         Index           =   4
         Left            =   2040
         TabIndex        =   6
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Specialization:"
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
         Index           =   9
         Left            =   4440
         TabIndex        =   42
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Qualfications:"
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
         Index           =   8
         Left            =   4440
         TabIndex        =   41
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Address:"
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
         Index           =   5
         Left            =   120
         TabIndex        =   40
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Home Phone:"
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
         Index           =   6
         Left            =   4440
         TabIndex        =   39
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Mobile Phone:"
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
         Index           =   7
         Left            =   4440
         TabIndex        =   38
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "First Name:"
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
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Last Name:"
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
         Left            =   4440
         TabIndex        =   31
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Sex:"
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
         Left            =   120
         TabIndex        =   30
         Top             =   885
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "NIC No:"
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
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   1815
      End
   End
   Begin VB.TextBox txtFields 
      Appearance      =   0  'Flat
      DataField       =   "Doctor_ID"
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Top             =   780
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Employee Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   120
      TabIndex        =   33
      Top             =   4680
      Width           =   8895
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Doctor_Basic_Sal"
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   15
         Tag             =   "Amt"
         Top             =   1440
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.ComboBox cmbDoctorType 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmAddDoctorDetails.frx":4831
         Left            =   2040
         List            =   "frmAddDoctorDetails.frx":483B
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   240
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Doctor_Notes"
         Height          =   885
         Index           =   13
         Left            =   2040
         TabIndex        =   13
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Doctor_Type"
         Height          =   285
         Index           =   10
         Left            =   2040
         TabIndex        =   11
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Doctor_VCharge"
         Height          =   285
         Index           =   11
         Left            =   6480
         TabIndex        =   12
         Tag             =   "Amt"
         Top             =   315
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "Doctor_CCharge"
         Height          =   285
         Index           =   12
         Left            =   6480
         TabIndex        =   14
         Tag             =   "Amt"
         Top             =   870
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Basic Salary:"
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
         Index           =   15
         Left            =   4440
         TabIndex        =   48
         Top             =   1440
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Notes:"
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
         Index           =   13
         Left            =   120
         TabIndex        =   37
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Doctor Type:"
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
         Index           =   10
         Left            =   120
         TabIndex        =   36
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Visiting Charge:"
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
         Index           =   11
         Left            =   4440
         TabIndex        =   35
         Top             =   315
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Channeling Charge:"
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
         Index           =   12
         Left            =   4440
         TabIndex        =   34
         Top             =   885
         Width           =   1815
      End
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "Doctor Details"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   14
      Left            =   3480
      TabIndex        =   47
      Top             =   120
      Width           =   2475
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Doctor ID:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   780
      Width           =   1815
   End
End
Attribute VB_Name = "frmAddDoctorDetails"
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




Private Sub cmbDoctorType_Click()

If cmbDoctorType = "Permanent Doctor" Then
    txtFields(14).Locked = False
    
End If
If cmbDoctorType = "Visiting Doctor" Then
    txtFields(14).Locked = True
    txtFields(14) = ""
End If


End Sub



Private Sub cmdViewAll_Click()
frmDoctorDetails.Show
End Sub

Private Sub Form_Load()


Call Functions.DisableMenu
Set adoPrimaryRS = New Recordset
adoPrimaryRS.Open "select Doctor_ID,Doctor_FName,Doctor_LName,Doctor_Sex,Doctor_NID,Doctor_Address,Doctor_HPhone,Doctor_MPhone,Doctor_Qualfication,Doctor_Specialization,Doctor_Type,Doctor_VCharge,Doctor_CCharge,Doctor_Notes,Doctor_Basic_Sal from Doctor_Details", cnPatients, adOpenStatic, adLockOptimistic

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
  Dim DocID As String
  Dim oText As TextBox
  Dim rsDocs As Recordset
  Set rsDocs = New ADODB.Recordset
  
  'On Error GoTo AddErr
  
        DocID = Functions.UID(6, "DocID_")
        rsDocs.Open "Select * from Doctor_Details", cnPatients, adOpenKeyset, adLockPessimistic
              
            While rsDocs.EOF = False
                If rsDocs(0) = DocID Then
                    DocID = Functions.UID(6, "DocID_")
                    flag = True
                    rsDocs.MoveFirst
                Else
                    flag = False
                End If
                rsDocs.MoveNext
            Wend
            
  
            rsDocs.Close
  cmbDoctorType_Click
  
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    
    For Each oText In Me.txtFields
        oText.Enabled = True
        oText.Locked = False
    Next
    cmbDoctorType.Visible = True
    cmbDoctorType.Text = cmbDoctorType.List(0)
    txtFields(10) = cmbDoctorType.Text
    
    txtFields(0).Text = DocID
    txtFields(0).Enabled = False
    txtFields(10).Enabled = False
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
  Dim oText As TextBox
  On Error GoTo EditErr
  

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  

  'Bind the text boxes to the data provider
    For Each oText In Me.txtFields
        oText.Enabled = True
        oText.Locked = False
    Next
    txtFields(0).Enabled = False
    txtFields(10).Enabled = False
    cmbDoctorType.Visible = True
    cmbDoctorType.Text = txtFields(10)
    cmbDoctorType_Click
  Exit Sub
    

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  Dim oText As TextBox
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
  
  
    For Each oText In Me.txtFields
        oText.Locked = True
    Next
    
    cmbDoctorType.Visible = False
    mbDataChanged = False
End Sub

Private Sub cmdUpdate_Click()
  Dim oText As TextBox
  On Error GoTo UpdateErr

  txtFields(10) = cmbDoctorType.Text
  txtFields(3) = cmbGender
  
  If txtFields(0) = "" Or txtFields(2) = "" Or txtFields(3) = "" Or txtFields(4) = "" Or txtFields(5) = "" Or txtFields(8) = "" Or txtFields(9) = "" Or txtFields(10) = "" Or txtFields(11) = "" Or txtFields(12) = "" Then
    MsgBox "Some of the required fields are missing", vbCritical, "Error Occured"
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
    oText.Enabled = False
    oText.Locked = True
  Next
  cmbDoctorType.Visible = False
  

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
  cmdViewAll.Visible = bVal
  
  cmbGender.Visible = Not bVal
  
  txtFields(14).Visible = Not bVal
  lblLabels(15).Visible = Not bVal
  


End Sub

