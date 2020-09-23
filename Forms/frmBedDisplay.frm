VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmBEDDisplay 
   BackColor       =   &H00FF8080&
   Caption         =   "Room Ward Bed Details"
   ClientHeight    =   9960
   ClientLeft      =   165
   ClientTop       =   195
   ClientWidth     =   12300
   Icon            =   "frmBedDisplay.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9960
   ScaleWidth      =   12300
   WindowState     =   2  'Maximized
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
      Height          =   855
      Left            =   10800
      Picture         =   "frmBedDisplay.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8760
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   8640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList TreeImages 
      Left            =   1680
      Top             =   8640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView OrgTree 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   12938
      _Version        =   393217
      LineStyle       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
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
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   7335
      Left            =   3240
      TabIndex        =   1
      Top             =   1080
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   12938
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      BackColorBkg    =   16744576
      GridColor       =   0
      HighLight       =   2
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      FormHeightDT    =   10470
      FormWidthDT     =   12420
      FormScaleHeightDT=   9960
      FormScaleWidthDT=   12300
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "BED / WARD DETAILS"
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
      Left            =   4200
      TabIndex        =   3
      Top             =   240
      Width           =   4440
   End
   Begin VB.Image TreeImage 
      Height          =   285
      Index           =   1
      Left            =   1320
      Picture         =   "frmBedDisplay.frx":0DCE
      Top             =   3720
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image TreeImage 
      Height          =   255
      Index           =   2
      Left            =   1320
      Picture         =   "frmBedDisplay.frx":18EF
      Top             =   4080
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image TreeImage 
      Height          =   285
      Index           =   3
      Left            =   1320
      Picture         =   "frmBedDisplay.frx":23BC
      Top             =   4440
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image TreeImage 
      Height          =   285
      Index           =   4
      Left            =   1680
      Picture         =   "frmBedDisplay.frx":2ED1
      Top             =   3720
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image TreeImage 
      Height          =   255
      Index           =   5
      Left            =   1680
      Picture         =   "frmBedDisplay.frx":3986
      Top             =   4080
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image TreeImage 
      Height          =   285
      Index           =   6
      Left            =   1680
      Picture         =   "frmBedDisplay.frx":4454
      Top             =   4440
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   1
      Left            =   1320
      Top             =   4800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   2
      Left            =   1320
      Top             =   5280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image IconImage 
      Height          =   480
      Index           =   3
      Left            =   1320
      Top             =   5760
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "frmBEDDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsWard As ADODB.Recordset
Dim rsBed As ADODB.Recordset
Dim rsPatientDe As ADODB.Recordset
Dim rsdata As ADODB.Recordset
'table name for the ward
'Ward_ID Ward_Name   Ward_Rate   Ward_Desc
Private Enum ObjectType
    otNone = 0
    otFactory = 1
    otGroup = 2
    otPerson = 3
    otFactory2 = 4
    otGroup2 = 5
    otPerson2 = 6
End Enum

Private SourceNode As Object
Private SourceType As ObjectType
Private TargetNode As Object
Private colheader As ColumnHeader
Private rslist As ADODB.Recordset
Dim LItem As ListItem
Dim selection As String
Dim getid As String

   



Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call Functions.DisableMenu

Dim colheader As ColumnHeader
    Dim intloopindex As Integer
    Dim Node1, Node2, Node3, Node4 As Node
  


Dim i As Integer
Dim name As String
Dim name1 As String
Dim ID As String
Dim factory As Node
Dim factory1 As Node
Dim group As Node
Dim person As Node

Set rsWard = New ADODB.Recordset
Set rsBed = New ADODB.Recordset
Set rsPatientDe = New ADODB.Recordset
Set rsdata = New ADODB.Recordset
' Load pictures into the ImageList.
    For i = 1 To 6
        TreeImages.ListImages.add , , TreeImage(i).Picture
    Next i
    
    ' Attach the TreeView to the ImageList.
    OrgTree.ImageList = TreeImages
    
    ' Create some nodes.
    rsWard.Open "select ward_name,Ward_ID from ward_details", cnPatients, adOpenDynamic, adLockPessimistic
    rsWard.MoveFirst
    Set factory = OrgTree.Nodes.add(, , "ward", "Ward Managment", otFactory, otFactory2)
    i = 0
    While Not rsWard.EOF
    name = rsWard!Ward_Name
    ID = rsWard!Ward_ID
    Set group = OrgTree.Nodes.add(factory, tvwChild, i & name & "", "" & name & "", otGroup, otGroup2)
    
    rsBed.Open "SELECT [Bed_ID] FROM [Bed_Details] where room_ward_id= '" & ID & "'", cnPatients, adOpenDynamic, adLockPessimistic
    
    If rsBed.RecordCount > 0 Then
    rsBed.MoveFirst
    While Not rsBed.EOF
    Debug.Print rsBed(0)
    name = rsBed(0)
    Set person = OrgTree.Nodes.add(group, tvwChild, i & name & "", "" & name & "", otPerson, otPerson2)
    rsBed.MoveNext
    i = i + 1
    Wend
    
    End If
   rsBed.Close
    rsWard.MoveNext
    i = i + 1
    Wend
    rsWard.Close
    
       
    Set factory = OrgTree.Nodes.add(, , "bed", "Room Managment", otFactory, otFactory2)
    rsWard.Open "SELECT [Room_ID], [Room_Type] FROM [Room_Details]", cnPatients, adOpenDynamic, adLockPessimistic
    rsWard.MoveFirst
    While Not rsWard.EOF
    i = i + 1
    name = rsWard!Room_Type
    ID = rsWard(0)
    Set group = OrgTree.Nodes.add(factory, tvwChild, i & name & "", "" & name & "", otGroup, otGroup2)
    Debug.Print ID
      
    rsBed.Open "select Bed_ID from bed_Details where room_ward_id= '" & ID & "'", cnPatients, adOpenDynamic, adLockPessimistic
    If rsBed.RecordCount > 0 Then
    rsBed.MoveFirst
    While Not rsBed.EOF
    Debug.Print rsBed(0)
    name = rsBed(0)
    Set person = OrgTree.Nodes.add(group, tvwChild, "hf " & name & "", "" & name & "", otPerson, otPerson2)
    rsBed.MoveNext
    Wend
    End If
    rsBed.Close
    
    rsWard.MoveNext
    Wend
    rsWard.Close
            
    
   
 
 'Framesettings.Visible = False
'cmdsettingshow.Visible = True
'cmdsettinghide.Visible = False
End Sub

Private Sub mnuclose_Click(Index As Integer)
Unload Me
End Sub



Private Sub Form_Unload(Cancel As Integer)
Call Functions.EnableMenu
End Sub

Private Sub OrgTree_NodeClick(ByVal Node As MSComctlLib.Node)
Dim icol As Integer
Dim getadmid As String
Dim avil As Boolean

Dim rec As Integer
Dim X As Integer
Dim clicked As String
Dim q As Boolean
Dim a As Collection


clicked = Node.Text
If clicked = "" Then
q = False
Else
q = True
End If

MSFlexGrid1.clear
Set rslist = New ADODB.Recordset
If clicked = "Ward Managment" Then

rslist.Open "select * from ward_details", cnPatients, adOpenDynamic, adLockPessimistic
rec = rslist.RecordCount

If rslist.EOF = False Then

With MSFlexGrid1
    .clear
    .Rows = 1
    .Cols = rslist.Fields.Count
   
  

    While Not rslist.EOF
        .Rows = .Rows + 1
        .Row = .Rows - 1


        For icol = 0 To rslist.Fields.Count - 1
            .Col = icol
            .Text = rslist(icol) & ""
        Next
        rslist.MoveNext
    Wend

   .TextMatrix(0, 0) = "Ward ID"
    .TextMatrix(0, 1) = "Ward name"
    .TextMatrix(0, 2) = "Ward Rate"
    .TextMatrix(0, 3) = "Ward Desc"
    
    .FixedRows = 1
    .RowHeight(0) = .RowHeight(1) * 1.5
    .ColWidth(1) = .ColWidth(0) * 2.5
    .ColWidth(2) = .ColWidth(0) * 1.75
    .ColWidth(3) = .ColWidth(0) * 1.5
    
End With

rslist.Close

End If


ElseIf clicked = "Room Managment" Then

rslist.Open "select * from Room_Details", cnPatients, adOpenDynamic, adLockPessimistic

If rslist.EOF = False Then

With MSFlexGrid1
    .clear
    .Rows = 1
    .Cols = rslist.Fields.Count
   
  

    While Not rslist.EOF
        .Rows = .Rows + 1
        .Row = .Rows - 1


        For icol = 0 To rslist.Fields.Count - 1
            .Col = icol
            .Text = rslist(icol) & ""
        Next
        rslist.MoveNext
    Wend

   .TextMatrix(0, 0) = "Room ID"
   .TextMatrix(0, 1) = "Room Type"
   .TextMatrix(0, 2) = "Room Rate"
   .TextMatrix(0, 3) = "Room Description"

    .FixedRows = 1
    .RowHeight(0) = .RowHeight(1) * 1.5
    .ColWidth(1) = .ColWidth(0) * 2.5
    .ColWidth(2) = .ColWidth(0) * 1.75
    .ColWidth(3) = .ColWidth(0) * 1.5
    
End With

rslist.Close


End If


Else

    rslist.Open "SELECT [Ward_ID], [Ward_Name] FROM [Ward_Details] where [Ward_Name]='" & clicked & "'", cnPatients, adOpenDynamic, adLockPessimistic

        If rslist.RecordCount > 0 Then
        
            getid = rslist(0)
                    
        End If

        If rslist.RecordCount = 0 Then

            rsBed.Open "SELECT [Room_ID], [Room_Type] FROM [Room_Details] where [Room_Type]='" & clicked & "'", cnPatients, adOpenDynamic, adLockPessimistic



                If rsBed.RecordCount > 0 Then
         
                     getid = rsBed(0)
        
                End If

            rsBed.Close
            rslist.Close

        End If





rsBed.Open "SELECT [Bed_ID], [Room_Ward_ID], [Available], [Admission_ID], [Bed_Desc] FROM [Bed_Details] where [Room_Ward_ID]='" & getid & "'", cnPatients, adOpenDynamic, adLockOptimistic
rec = rsBed.RecordCount

If rsBed.EOF = False Then

    With MSFlexGrid1
        .clear
         .Rows = 1
         .Cols = rsBed.Fields.Count
        
  

        While Not rsBed.EOF
            .Rows = .Rows + 1
            .Row = .Rows - 1


        For icol = 0 To rsBed.Fields.Count - 1
            .Col = icol
            .Text = rsBed(icol) & ""
        Next
        rsBed.MoveNext
         Wend

    .TextMatrix(0, 0) = "Bed ID"
    .TextMatrix(0, 1) = "Room Ward ID"
    .TextMatrix(0, 2) = "Availability"
    .TextMatrix(0, 3) = "Admission ID"
    .TextMatrix(0, 4) = "Bed Discription"
    
    .FixedRows = 1
    .RowHeight(0) = .RowHeight(1) * 1.5
    .ColWidth(1) = .ColWidth(0) * 2.5
    .ColWidth(2) = .ColWidth(0) * 1.75
    .ColWidth(3) = .ColWidth(0) * 1.5
    .ColWidth(4) = .ColWidth(0) * 1.5
    
End With



End If
rsBed.Close


rsPatientDe.Open "SELECT [Bed_ID], [Room_Ward_ID], [Available], [Admission_ID], [Bed_Desc] FROM [Bed_Details] where [Bed_ID]='" & clicked & "'", cnPatients, adOpenDynamic, adLockOptimistic

If rsPatientDe.RecordCount = 1 Then

'getadmid = rsPatientDe(3)
avil = rsPatientDe(2)

    If avil = False Then
        getadmid = rsPatientDe(3)
        Debug.Print getadmid
    
    End If

End If
rsPatientDe.Close


rsdata.Open "SELECT [Admission_ID], [Patient_ID], [Guardian_ID], [Ref_Doctor], [Admission_Date], [Admission_Time] FROM [Admission_Details] where [Admission_ID]='" & getadmid & "'", cnPatients, adOpenDynamic, adLockOptimistic

rec = rsdata.RecordCount

If rsdata.EOF = False Then

With MSFlexGrid1
    .clear
    .Rows = 1
    .Cols = rsdata.Fields.Count
   
  

    While Not rsdata.EOF
        .Rows = .Rows + 1
        .Row = .Rows - 1


        For icol = 0 To rsdata.Fields.Count - 1
            .Col = icol
            .Text = rsdata(icol) & ""
        Next
        rsdata.MoveNext
    Wend

    .TextMatrix(0, 0) = "Admission ID"
    .TextMatrix(0, 1) = "Patient ID"
    .TextMatrix(0, 2) = "Guardian ID"
    .TextMatrix(0, 3) = "Ref Doctor"
    .TextMatrix(0, 4) = "Admission Date"
    .TextMatrix(0, 5) = "Admission Time"
    
    .FixedRows = 1
    .RowHeight(0) = .RowHeight(1) * 1.5
    .ColWidth(1) = .ColWidth(0) * 2.5
    .ColWidth(2) = .ColWidth(0) * 1.75
    .ColWidth(3) = .ColWidth(0) * 1.5
    .ColWidth(4) = .ColWidth(0) * 1.5
    .ColWidth(5) = .ColWidth(0) * 1.5
    
End With

End If
rsdata.Close

End If
Functions.SizeColumns MSFlexGrid1, Me
Functions.SizeColumnHeaders MSFlexGrid1, Me

End Sub





