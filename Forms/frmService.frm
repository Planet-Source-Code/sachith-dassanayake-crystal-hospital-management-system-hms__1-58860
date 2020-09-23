VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmService 
   BackColor       =   &H00FF8080&
   Caption         =   "Hospital Services Details"
   ClientHeight    =   10800
   ClientLeft      =   2610
   ClientTop       =   2340
   ClientWidth     =   11340
   Icon            =   "frmService.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10800
   ScaleWidth      =   11340
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Controls"
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
      Height          =   1335
      Left            =   2880
      TabIndex        =   6
      Top             =   8400
      Width           =   5775
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   795
         Left            =   4320
         Picture         =   "frmService.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Search"
         Height          =   795
         Left            =   600
         Picture         =   "frmService.frx":0DCE
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   795
         Left            =   2400
         Picture         =   "frmService.frx":1251
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.ComboBox cmbSearch 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   10080
      Width           =   2535
   End
   Begin VB.TextBox txtSearchText 
      Height          =   315
      Left            =   7080
      TabIndex        =   1
      Top             =   10080
      Width           =   2535
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6855
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   12091
      View            =   3
      Arrange         =   2
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   0
      BackColor       =   -2147483624
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Service ID"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Medical Service"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Avarage Duration"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Charge"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Notes"
         Object.Width           =   2540
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
      FormHeightDT    =   11310
      FormWidthDT     =   11460
      FormScaleHeightDT=   10800
      FormScaleWidthDT=   11340
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Search For"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   10080
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "Search Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   10080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "VIEW HOSPITAL SERVICES DETAILS"
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
      Left            =   1725
      TabIndex        =   0
      Top             =   240
      Width           =   7305
   End
End
Attribute VB_Name = "frmService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strCol As Variant


Private Sub cmdAddnew_Click()
frmAddOutPatientDetails.Show
End Sub





Private Sub cmbSearch_Click()
cmdFind_Click
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()

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
        strSQl = "SELECT * FROM Services WHERE "
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
    
        
        
        'LItem.SubItems(1) = rsFind(1)
        'LItem.SubItems(2) = rsFind(2)
        'LItem.SubItems(3) = rsFind(3)
        'LItem.SubItems(4) = rsFind(4)
        'LItem.SubItems(5) = rsFind(5)
        
        rsFind.MoveNext
        Wend
    End If
 
 
 'show number of records found
    Me.Caption = CStr(rsFind.RecordCount) & " records found"
    
 'close the recordset
    rsFind.Close
    
    
End Sub

Private Sub cmdSearch_Click()
Dim LItem As ListItem

FindItem = InputBox("Enter Patient ID", "Find Patient Details", "DocID_")

If Not FindItem = "" Then
Set LItem = ListView1.FindItem(FindItem, lvwText, lvwSubItem)
If LItem Is Nothing Then
    NotFound = True
End If
If NotFound Then
    MsgBox "Item not found", vbInformation, "Search Result"
Else
    LItem.EnsureVisible
    LItem.Selected = True
End If
End If


End Sub






Private Sub Command1_Click()
txtSearchText = ""
Form_Load
End Sub

Private Sub Form_Activate()
Call Functions.DisableMenu
Me.WindowState = vbMaximized
txtSearchText.SetFocus
End Sub

Private Sub Form_Load()


Me.WindowState = vbMaximized



'On Error Resume Next
Call Functions.DisableMenu


Dim LItem As ListItem
Dim i As Integer


Dim rsPatientID As Recordset
Set rsPatientID = New ADODB.Recordset

rsPatientID.Open "select * from Services", cnPatients, adOpenDynamic, adLockPessimistic


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

Private Sub txtSearchText_Change()
cmdFind_Click
End Sub



