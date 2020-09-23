VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form FrmCustomers 
   BackColor       =   &H00FF8080&
   Caption         =   "Pharmacy - Customer Details"
   ClientHeight    =   9195
   ClientLeft      =   1110
   ClientTop       =   450
   ClientWidth     =   13620
   Icon            =   "FrmCustomers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   13620
   WindowState     =   2  'Maximized
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
      Left            =   2280
      TabIndex        =   40
      Top             =   5400
      Width           =   8895
      Begin VB.CommandButton cmdLast 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   7200
         Picture         =   "FrmCustomers.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   6480
         Picture         =   "FrmCustomers.frx":1317
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   2025
         Picture         =   "FrmCustomers.frx":17F2
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdFirst 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   1320
         Picture         =   "FrmCustomers.frx":1CD3
         Style           =   1  'Graphical
         TabIndex        =   41
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
         Height          =   285
         Left            =   2880
         TabIndex        =   45
         Top             =   360
         Width           =   3360
      End
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
      Height          =   2535
      Left            =   2280
      TabIndex        =   31
      Top             =   6480
      Width           =   8895
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
         Left            =   2160
         Picture         =   "FrmCustomers.frx":21A9
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   360
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
         Left            =   3480
         Picture         =   "FrmCustomers.frx":265C
         Style           =   1  'Graphical
         TabIndex        =   38
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
         Left            =   4800
         Picture         =   "FrmCustomers.frx":2B33
         Style           =   1  'Graphical
         TabIndex        =   37
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
         Left            =   6120
         Picture         =   "FrmCustomers.frx":2FF6
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   360
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
         Left            =   4800
         Picture         =   "FrmCustomers.frx":349C
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Update"
         Height          =   780
         Left            =   2160
         Picture         =   "FrmCustomers.frx":39A8
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
         Left            =   3480
         Picture         =   "FrmCustomers.frx":3E56
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdViewAll 
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
         Left            =   3480
         Picture         =   "FrmCustomers.frx":435A
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.ComboBox cmbContactTitle 
      Height          =   315
      ItemData        =   "FrmCustomers.frx":4815
      Left            =   3240
      List            =   "FrmCustomers.frx":4825
      TabIndex        =   30
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ContactTitle"
      Height          =   285
      Index           =   4
      Left            =   3240
      TabIndex        =   29
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Notes"
      Height          =   990
      Index           =   14
      Left            =   9240
      TabIndex        =   28
      Tag             =   "Chr"
      Top             =   3465
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PostalCode"
      Height          =   285
      Index           =   13
      Left            =   9240
      TabIndex        =   26
      Tag             =   "Num"
      Top             =   3015
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Country/Region"
      Height          =   285
      Index           =   12
      Left            =   9240
      TabIndex        =   24
      Tag             =   "Chr"
      Top             =   2580
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "City"
      Height          =   285
      Index           =   11
      Left            =   9240
      TabIndex        =   22
      Tag             =   "Chr"
      Top             =   2145
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "BillingAddress"
      Height          =   885
      Index           =   10
      Left            =   9240
      TabIndex        =   20
      Top             =   1095
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "StateOrProvince"
      Height          =   285
      Index           =   9
      Left            =   9240
      TabIndex        =   18
      Tag             =   "Chr"
      Top             =   660
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Extension"
      Height          =   285
      Index           =   8
      Left            =   3240
      MaxLength       =   5
      TabIndex        =   16
      Tag             =   "Num"
      Top             =   4185
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      DataField       =   "FaxNumber"
      Height          =   285
      Index           =   7
      Left            =   3240
      MaxLength       =   10
      TabIndex        =   14
      Tag             =   "Num"
      Top             =   3735
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "PhoneNumber"
      Height          =   285
      Index           =   6
      Left            =   3240
      MaxLength       =   10
      TabIndex        =   12
      Tag             =   "Num"
      Top             =   3300
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      DataField       =   "EmailAddress"
      Height          =   285
      Index           =   5
      Left            =   3240
      TabIndex        =   10
      Top             =   2865
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ContactLastName"
      Height          =   285
      Index           =   3
      Left            =   3240
      TabIndex        =   7
      Tag             =   "Chr"
      Top             =   1980
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ContactFirstName"
      Height          =   285
      Index           =   2
      Left            =   3240
      TabIndex        =   5
      Tag             =   "Chr"
      Top             =   1545
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CompanyName"
      Height          =   285
      Index           =   1
      Left            =   3240
      TabIndex        =   3
      Tag             =   "Chr"
      Top             =   1095
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      DataField       =   "CustomerID"
      Height          =   285
      Index           =   0
      Left            =   3240
      TabIndex        =   1
      Top             =   660
      Width           =   1695
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
      FormHeightDT    =   9705
      FormWidthDT     =   13740
      FormScaleHeightDT=   9195
      FormScaleWidthDT=   13620
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Height          =   4815
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   13215
   End
   Begin VB.Label lblLabels 
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
      Height          =   255
      Index           =   14
      Left            =   7320
      TabIndex        =   27
      Top             =   3465
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "PostalCode:"
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
      Index           =   13
      Left            =   7320
      TabIndex        =   25
      Top             =   3015
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Country/Region:"
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
      Index           =   12
      Left            =   7320
      TabIndex        =   23
      Top             =   2580
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "City:"
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
      Index           =   11
      Left            =   7320
      TabIndex        =   21
      Top             =   2145
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "BillingAddress:"
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
      Index           =   10
      Left            =   7320
      TabIndex        =   19
      Top             =   1095
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "StateOrProvince:"
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
      Index           =   9
      Left            =   7320
      TabIndex        =   17
      Top             =   660
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "Extension:"
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
      Index           =   8
      Left            =   1320
      TabIndex        =   15
      Top             =   4185
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "FaxNumber:"
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
      Index           =   7
      Left            =   1320
      TabIndex        =   13
      Top             =   3735
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "PhoneNumber:"
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
      Left            =   1320
      TabIndex        =   11
      Top             =   3300
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "EmailAddress:"
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
      Left            =   1320
      TabIndex        =   9
      Top             =   2865
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "ContactTitle:"
      DataField       =   "ContactTitle:"
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
      Left            =   1320
      TabIndex        =   8
      Top             =   2415
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "ContactLastName:"
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
      Left            =   1320
      TabIndex        =   6
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "ContactFirstName:"
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
      Left            =   1320
      TabIndex        =   4
      Top             =   1545
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "CompanyName:"
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
      Left            =   1320
      TabIndex        =   2
      Top             =   1095
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      Caption         =   "CustomerID:"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   660
      Width           =   1815
   End
End
Attribute VB_Name = "FrmCustomers"
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

Private Sub cmdViewAll_Click()
FrmVCustomers.Show

End Sub

Private Sub Form_Load()

  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select * from Customers", cnPatients, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
    oText.Enabled = True
    oText.Locked = True
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
    
    Dim rsAddCustomers As New Recordset
    Dim CID As String
    Set rsAddCustomers = New ADODB.Recordset
  
    CID = Functions.UID(6, "CID_")
    rsAddCustomers.Open " Select * from Customers", cnPatients, adOpenKeyset, adLockPessimistic
    While rsAddCustomers.EOF = False
        If rsAddCustomers(0) = CID Then
            CID = Functions.UID(6, "CID_")
            rsAddCustomers.MoveFirst
        Else
           
        End If
    rsAddCustomers.MoveNext
    Wend
    rsAddCustomers.Close
    
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    oText.Locked = False
  Next
    
  txtFields(0).Locked = True
  
  
  
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    txtFields(0).Text = CID
    
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
  txtFields(4) = cmbContactTitle
  
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
  cmbContactTitle.Visible = Not bVal
  cmdViewAll.Visible = bVal
  
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 22 Then KeyAscii = 0: Exit Sub
    KeyAscii = DataEntryValidation(KeyAscii, ActiveControl.Tag)
End Sub

