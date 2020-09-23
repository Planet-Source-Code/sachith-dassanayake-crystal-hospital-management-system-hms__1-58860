VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmViewServiceAppointments 
   BackColor       =   &H00FF8080&
   Caption         =   "View Hospital Service Appointments"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12465
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmViewServiceAppointments.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   12465
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Height          =   1215
      Left            =   360
      TabIndex        =   10
      Top             =   1440
      Width           =   11775
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Display"
         Height          =   840
         Left            =   10320
         Picture         =   "frmViewServiceAppointments.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkDoc 
         BackColor       =   &H00FF8080&
         Caption         =   "Service ID"
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
         TabIndex        =   12
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cmbService 
         Height          =   315
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpDateTo 
         Height          =   375
         Left            =   4680
         TabIndex        =   13
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   46137345
         CurrentDate     =   38350
      End
      Begin MSComCtl2.DTPicker dtpDateFrom 
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   46137345
         CurrentDate     =   38350
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
         TabIndex        =   16
         Top             =   480
         Width           =   735
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
         TabIndex        =   15
         Top             =   480
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Height          =   1335
      Left            =   360
      TabIndex        =   5
      Top             =   6240
      Width           =   7455
      Begin VB.TextBox txtSearchText 
         Height          =   315
         Left            =   5760
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cmbSearch 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   2535
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
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   1335
      Left            =   8160
      TabIndex        =   2
      Top             =   6240
      Width           =   3975
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   975
         Left            =   2160
         Picture         =   "frmViewServiceAppointments.frx":5C9D
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Refresh"
         Height          =   975
         Left            =   720
         Picture         =   "frmViewServiceAppointments.frx":61A1
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   360
      TabIndex        =   1
      Top             =   2880
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5530
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Appointment ID"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Patient ID"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Hospital Service ID"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Appointment Date"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Appointment Time"
         Object.Width           =   3528
      EndProperty
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   3135
      Left            =   7800
      TabIndex        =   0
      Top             =   2880
      Width           =   4335
      _Version        =   524288
      _ExtentX        =   7646
      _ExtentY        =   5530
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2004
      Month           =   12
      Day             =   29
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      FormHeightDT    =   8700
      FormWidthDT     =   12585
      FormScaleHeightDT=   8190
      FormScaleWidthDT=   12465
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "VIEW SERVICE APPOINTMENTS"
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
      Left            =   3210
      TabIndex        =   18
      Top             =   480
      Width           =   6345
   End
End
Attribute VB_Name = "frmViewServiceAppointments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strCol As Variant

Private Sub Calendar1_Click()

Dim DateClicked As Date
Dim rsAppointment As Recordset
Dim LItem As ListItem
Dim LHeader As ColumnHeader
Dim SQL As String

Set rsAppointment = New ADODB.Recordset

DateClicked = Calendar1.Value
SQL = "SELECT * From Service_Appointment, Patient_Details WHERE Service_Appointment.Appointment_Date =#" & SQLDate(DateClicked) & "#" & " and Service_Appointment.Patient_ID=Patient_Details.Patient_ID ORDER BY Appointment_Time" & ";"
    
Set rsAppointment = New ADODB.Recordset

ListView1.ListItems.clear
ListView1.ColumnHeaders.clear

rsAppointment.Open SQL, cnPatients, adOpenDynamic, adLockPessimistic

LWidth = ListView1.Width - 5 * Screen.TwipsPerPixelX
Set LHeader = ListView1.ColumnHeaders.add(1, , "Appointment ID", 2000)
Set LHeader = ListView1.ColumnHeaders.add(2, , "Patient First Name", 2000)
Set LHeader = ListView1.ColumnHeaders.add(3, , "Patient Last Name", 2000)
Set LHeader = ListView1.ColumnHeaders.add(4, , "Address", 2000)
Set LHeader = ListView1.ColumnHeaders.add(5, , "Telephone", 2000)
Set LHeader = ListView1.ColumnHeaders.add(6, , "Service ID", 2000)
Set LHeader = ListView1.ColumnHeaders.add(7, , "Appointment Time", 2000)


Dim rsDocName As Recordset
Set rsDocName = New ADODB.Recordset


If Not rsAppointment.EOF Then ' If results found
   

            While Not rsAppointment.EOF
                'rsDocName.Open "select Doctor_Fname,Doctor_LName from Doctor_Details where Doctor_ID='" & rsAppointment![Doctor_ID] & "'", cnPatients, adOpenDynamic, adLockPessimistic
        
                Set LItem = ListView1.ListItems.add(, , rsAppointment![Appointment_ID])
                
                If rsAppointment![First_Name] <> "" Then
                    LItem.SubItems(1) = rsAppointment![First_Name]
                End If
                If rsAppointment![Last_Name] <> "" Then
                    LItem.SubItems(2) = rsAppointment![Last_Name]
                End If
                If rsAppointment![address] <> "" Then
                    LItem.SubItems(3) = rsAppointment![address]
                End If
                If rsAppointment![Telephone] <> "" Then
                    LItem.SubItems(4) = rsAppointment![Telephone]
                End If
                If rsAppointment![Hospital_SErvice_ID] <> "" Then
                    LItem.SubItems(5) = rsAppointment![Hospital_SErvice_ID]
                End If
                If rsAppointment![Appointment_Time] <> "" Then
                    LItem.SubItems(6) = rsAppointment![Appointment_Time]
                End If
                
                rsAppointment.MoveNext
                'rsDocName.Close

            Wend


Else ' If no Results Found

End If

rsAppointment.Close




















End Sub



Private Sub cmbSearch_Click()
txtSearchText_Change
End Sub

Private Sub Command1_Click()

If dtpDateFrom > dtpDateTo Then
    MsgBox "The (From) date has to be less than the (To) Date", vbCritical
    Exit Sub
End If



Dim LItem As ListItem
Dim i As Integer
Dim SQL As String

If chkDoc.Value = 0 Then
    SQL = "select * from Service_Appointment where Appointment_Date between #" & SQLDate(dtpDateFrom) & "#  AND #" & SQLDate(dtpDateTo) & "#"
ElseIf chkDoc.Value = 1 Then
    SQL = "select * from Service_Appointment where Hospital_Service_ID='" & cmbService & "' and  Appointment_Date between #" & SQLDate(dtpDateFrom) & "#  AND #" & SQLDate(dtpDateTo) & "#"
End If

Dim rsSerAppointments As Recordset
Set rsSerAppointments = New ADODB.Recordset


rsSerAppointments.Open SQL, cnPatients, adOpenDynamic, adLockPessimistic


For i = 0 To rsSerAppointments.Fields.Count - 1 Step 1
    cmbSearch.AddItem rsSerAppointments(i).name, i
Next i

ListView1.ListItems.clear
ListView1.ColumnHeaders.clear

Set LHeader = ListView1.ColumnHeaders.add(1, , "Appointment ID", 2000)
Set LHeader = ListView1.ColumnHeaders.add(2, , "Patient ID", 2000)
Set LHeader = ListView1.ColumnHeaders.add(3, , "Service ID", 2000)
Set LHeader = ListView1.ColumnHeaders.add(4, , "Appointment Date", 2000)
Set LHeader = ListView1.ColumnHeaders.add(5, , "Apointment Time", 2000)


ListView1.ListItems.clear
While rsSerAppointments.EOF = False
        Set LItem = ListView1.ListItems.add(, , rsSerAppointments(0))
        LItem.SubItems(1) = rsSerAppointments(1)
        LItem.SubItems(2) = rsSerAppointments(2)
        LItem.SubItems(3) = Format(rsSerAppointments(3), "short Date")
        LItem.SubItems(4) = Format(rsSerAppointments(4), "short time")
        
rsSerAppointments.MoveNext
Wend
rsSerAppointments.Close



















End Sub

Private Sub Command3_Click()
Form_Load
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call Functions.DisableMenu
Dim LItem As ListItem
Dim i As Integer


Dim rsSerAppointments As Recordset
Set rsSerAppointments = New ADODB.Recordset

rsSerAppointments.Open "select * from Service_Appointment order by Appointment_Date,Appointment_Time", cnPatients, adOpenDynamic, adLockPessimistic


For i = 0 To rsSerAppointments.Fields.Count - 1 Step 1
    cmbSearch.AddItem rsSerAppointments(i).name, i
Next i

ListView1.ListItems.clear
ListView1.ColumnHeaders.clear

Set LHeader = ListView1.ColumnHeaders.add(1, , "Appointment ID", 2000)
Set LHeader = ListView1.ColumnHeaders.add(2, , "Patient ID", 2000)
Set LHeader = ListView1.ColumnHeaders.add(3, , "Medical Service ID", 3000)
Set LHeader = ListView1.ColumnHeaders.add(4, , "Appointment Date", 2500)
Set LHeader = ListView1.ColumnHeaders.add(5, , "Apointment Time", 2500)


ListView1.ListItems.clear
While rsSerAppointments.EOF = False
        Set LItem = ListView1.ListItems.add(, , rsSerAppointments(0))
        LItem.SubItems(1) = rsSerAppointments(1)
        LItem.SubItems(2) = rsSerAppointments(2)
        LItem.SubItems(3) = Format(rsSerAppointments(3), "short Date")
        LItem.SubItems(4) = Format(rsSerAppointments(4), "short time")
        'LItem.SubItems(5) = rsSerAppointments(5)
        
rsSerAppointments.MoveNext
Wend
rsSerAppointments.Close
cmbSearch.Text = cmbSearch.List(0)


dtpDateFrom = Date
dtpDateTo = Date

Dim rsAddServices As Recordset
Set rsAddServices = New ADODB.Recordset
rsAddServices.Open "Select * from Services", cnPatients, adOpenDynamic, adLockReadOnly

While rsAddServices.EOF = False
cmbService.AddItem rsAddServices(0)
rsAddServices.MoveNext
Wend

rsAddServices.Close





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

Dim rsFind As Recordset
Dim strSQl As String
Dim SQL As String
Dim LItem As ListItem
Dim LHeader As ColumnHeader

'if there is nothing to search for then exit
If txtSearchText = "" Then
    Exit Sub
End If


ListView1.ListItems.clear
ListView1.ColumnHeaders.clear

Set rsFind = New ADODB.Recordset

Set LHeader = ListView1.ColumnHeaders.add(1, , "Appointment ID", 2000)
Set LHeader = ListView1.ColumnHeaders.add(2, , "Patient ID", 2000)
Set LHeader = ListView1.ColumnHeaders.add(3, , "Service ID", 2000)
Set LHeader = ListView1.ColumnHeaders.add(4, , "Appointment Date", 2000)
Set LHeader = ListView1.ColumnHeaders.add(5, , "Apointment Time", 2000)


'make the search
        strSQl = "SELECT * FROM Service_Appointment WHERE "
        strSQl = strSQl & cmbSearch & " Like " & "'%" & txtSearchText & "%'"

        'SQL = strSQl & " WHERE language LIKE '*" & Text1.Text & "*'"
        'strSQl = strSQl & SQL
        Debug.Print strSQl

        
'show the found records
    rsFind.Open strSQl, cnPatients, adOpenDynamic, adLockPessimistic
    
    
    Debug.Print rsFind.RecordCount
    Debug.Print rsFind.Fields.Count
    
    If Not (rsFind.BOF And rsFind.EOF) Then
        While rsFind.EOF = False
        Set LItem = ListView1.ListItems.add(, , rsFind(0))
        LItem.SubItems(1) = rsFind(1)
        LItem.SubItems(2) = rsFind(2)
        LItem.SubItems(3) = Format(rsFind(3), "short date")
        LItem.SubItems(4) = Format(rsFind(4), "long time")
        'LItem.SubItems(5) = rsFind(5)
        
        rsFind.MoveNext
        Wend
    End If
 
 
 'show number of records found
    Me.Caption = CStr(rsFind.RecordCount) & " records found"
    
 'close the recordset
    rsFind.Close




End Sub

