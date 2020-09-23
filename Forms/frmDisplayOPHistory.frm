VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmDisplayOPHistory 
   BackColor       =   &H00FF8080&
   Caption         =   "Display Out Patient History"
   ClientHeight    =   9525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12600
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDisplayOPHistory.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9525
   ScaleWidth      =   12600
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   1335
      Left            =   8040
      TabIndex        =   13
      Top             =   7560
      Width           =   4095
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Refresh"
         Height          =   975
         Left            =   720
         Picture         =   "frmDisplayOPHistory.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   975
         Left            =   2040
         Picture         =   "frmDisplayOPHistory.frx":0D70
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Height          =   1335
      Left            =   240
      TabIndex        =   8
      Top             =   7560
      Width           =   7455
      Begin VB.ComboBox cmbSearch 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtSearchText 
         Height          =   315
         Left            =   5760
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF8080&
         Caption         =   "Search For"
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
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF8080&
         Caption         =   "Search Text"
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
         Left            =   4320
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   1440
      Width           =   11775
      Begin VB.CommandButton cmdPatientID 
         Caption         =   "..."
         Height          =   255
         Left            =   9720
         TabIndex        =   18
         Top             =   480
         Width           =   495
      End
      Begin VB.ComboBox cmbPatient 
         Height          =   315
         Left            =   8040
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox chkDoc 
         BackColor       =   &H00FF8080&
         Caption         =   "Patient ID"
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
         Height          =   375
         Left            =   6600
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdDisplay 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Display"
         Height          =   840
         Left            =   10320
         Picture         =   "frmDisplayOPHistory.frx":1274
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpDateTo 
         Height          =   375
         Left            =   4680
         TabIndex        =   4
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20578305
         CurrentDate     =   38350
      End
      Begin MSComCtl2.DTPicker dtpDateFrom 
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20578305
         CurrentDate     =   38350
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "To"
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
         Left            =   3840
         TabIndex        =   7
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "From"
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
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   735
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4455
      Left            =   360
      TabIndex        =   16
      Top             =   2880
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7858
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "OP History ID"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Patient ID"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Doctor ID"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Time"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Description"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Prescription"
         Object.Width           =   6174
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
      FormHeightDT    =   10035
      FormWidthDT     =   12720
      FormScaleHeightDT=   9525
      FormScaleWidthDT=   12600
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "VIEW OUT PATIENT HISTORY"
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
      Left            =   3420
      TabIndex        =   17
      Top             =   480
      Width           =   5925
   End
End
Attribute VB_Name = "frmDisplayOPHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDisplay_Click()

If dtpDateFrom > dtpDateTo Then
    MsgBox "The (From) date has to be less than the (To) Date", vbCritical
    Exit Sub
End If

Dim LItem As ListItem
Dim i As Integer
Dim SQL As String

If chkDoc.Value = 0 Then
    SQL = "select * from OutPatient_Treatments where Date between #" & SQLDate(dtpDateFrom) & "#  AND #" & SQLDate(dtpDateTo) & "#"
ElseIf chkDoc.Value = 1 Then
    SQL = "select * from OutPatient_Treatments where Patient_ID='" & cmbPatient & "' and  Date between #" & SQLDate(dtpDateFrom) & "#  AND #" & SQLDate(dtpDateTo) & "#"
End If

Dim rsOPHistory As Recordset
Set rsOPHistory = New ADODB.Recordset


rsOPHistory.Open SQL, cnPatients, adOpenDynamic, adLockPessimistic


For i = 0 To rsOPHistory.Fields.Count - 1 Step 1
    cmbSearch.AddItem rsOPHistory(i).name, i
Next i

ListView1.ListItems.clear

'While rsOPHistory.EOF = False
 '       Set LItem = ListView1.ListItems.add(, , rsOPHistory(0))
  '      LItem.SubItems(1) = rsOPHistory(1)
   '     LItem.SubItems(2) = rsOPHistory(2)
    '    LItem.SubItems(3) = Format(rsOPHistory(3), "short Date")
     '   LItem.SubItems(4) = Format(rsOPHistory(4), "short time")
      '  LItem.SubItems(5) = rsDocAppointments(5)
        
'rsDocAppointments.MoveNext
'Wend

While rsOPHistory.EOF = False
        Set LItem = ListView1.ListItems.add(, , rsOPHistory(0))
        
        For j = 1 To rsOPHistory.Fields.Count - 1 Step 1
            If rsOPHistory(j) <> "" Then
                LItem.SubItems(j) = rsOPHistory(j)
            End If
        Next j
    
        
rsOPHistory.MoveNext
Wend



rsOPHistory.Close






















End Sub

Private Sub cmdPatientID_Click()
frmDisplayOutPatient.Show
End Sub

Private Sub cmdRefresh_Click()
txtSearchText = ""
Form_Load
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
Call Functions.DisableMenu
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Functions.EnableMenu
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

If strCol <> ColumnHeader Then
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = ColumnHeader.Index - 1
    strCol = ColumnHeader
Else
    ListView1.SortOrder = lvwDescending
    ListView1.SortKey = ColumnHeader.Index - 1
    strCol = ""
End If


End Sub

Private Sub Form_Load()

Call Functions.DisableMenu
Dim LItem As ListItem
Dim i As Integer

dtpDateFrom = Date
dtpDateTo = Date + 1

Dim rsPatientID As Recordset
Set rsPatientID = New ADODB.Recordset

rsPatientID.Open "select * from OutPatient_Treatments", cnPatients, adOpenDynamic, adLockPessimistic
cmbSearch.clear

For i = 0 To rsPatientID.Fields.Count - 1 Step 1
    cmbSearch.AddItem rsPatientID(i).name, i
Next i


If rsPatientID.EOF = False Then
    rsPatientID.MoveFirst
Else
    Exit Sub
End If

ListView1.ListItems.clear

While rsPatientID.EOF = False
        Set LItem = ListView1.ListItems.add(, , rsPatientID(0))
        
        For j = 1 To rsPatientID.Fields.Count - 1 Step 1
            If rsPatientID(j) <> "" Then
                LItem.SubItems(j) = rsPatientID(j)
            End If
        Next j
    
        
rsPatientID.MoveNext
Wend
rsPatientID.Close
cmbSearch.Text = cmbSearch.List(0)


cmbPatient.clear
Dim rsAddPat As Recordset
Set rsAddPat = New ADODB.Recordset
rsAddPat.Open "Select * from Patient_Details", cnPatients, adOpenDynamic, adLockReadOnly

While rsAddPat.EOF = False
cmbPatient.AddItem rsAddPat(0)
rsAddPat.MoveNext
Wend

rsAddPat.Close



End Sub

Private Sub txtSearchText_Change()
Dim rsFind As Recordset
Dim strSQl As String
Dim SQL As String
Dim LItem As ListItem

'if there is nothing to search for then exit
If txtSearchText = "" Then
    Exit Sub
End If

ListView1.ListItems.clear

Set rsFind = New ADODB.Recordset



'make the search
        strSQl = "SELECT * FROM OutPatient_Treatments WHERE "
        strSQl = strSQl & cmbSearch & " Like " & "'%" & txtSearchText & "%'"

   
        Debug.Print strSQl
        
'show the found records
    rsFind.Open strSQl, cnPatients, adOpenDynamic, adLockPessimistic
    
    
    Debug.Print rsFind.RecordCount
    Debug.Print rsFind.Fields.Count
    
    If Not (rsFind.BOF And rsFind.EOF) Then
        While rsFind.EOF = False
        Set LItem = ListView1.ListItems.add(, , rsFind(0))
        
           
        For j = 1 To rsFind.Fields.Count - 1 Step 1
            If rsFind(j) <> "" Then
                LItem.SubItems(j) = rsFind(j)
            End If
        Next j
            
        rsFind.MoveNext
        Wend
    End If
 
 
 'show number of records found
    Me.Caption = CStr(rsFind.RecordCount) & " records found"
    
 'close the recordset
    rsFind.Close
    
End Sub
