VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frm_employee_view 
   BackColor       =   &H00FF8080&
   Caption         =   "Employee Table View"
   ClientHeight    =   7740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14940
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_employee_view.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   14940
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   13320
      Picture         =   "frm_employee_view.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   8070
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16761024
      BackColorBkg    =   16744576
      GridColor       =   0
      AllowBigSelection=   0   'False
      HighLight       =   2
      GridLinesFixed  =   1
      AllowUserResizing=   1
      Appearance      =   0
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
      FormHeightDT    =   8250
      FormWidthDT     =   15060
      FormScaleHeightDT=   7740
      FormScaleWidthDT=   14940
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "VIEW EMPLOYEE DETAILS"
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
      Left            =   5040
      TabIndex        =   2
      Top             =   240
      Width           =   5220
   End
End
Attribute VB_Name = "frm_employee_view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3CBB305B037F"
Dim rsDocs As New Recordset
'##ModelId=3CBB305B0380
Private strst As String

Private Sub cmdClose_Click()
Unload Me
End Sub

'##ModelId=3CBB305B0389
Private Sub Form_Load()


Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset

'con.CursorLocation = adUseClient
'con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\pay\Pay.mdb;persist security info=false"
strst = "del"


Set rsDocs = New ADODB.Recordset


'create sql statement


rsDocs.Open "select * from employee  ", cnPatients, adOpenDynamic, adLockOptimistic



If rsDocs.EOF = False Then

With MSFlexGrid1
    .clear
    .Rows = 1
    .Cols = rsDocs.Fields.Count
   
  

    While Not rsDocs.EOF
        .Rows = .Rows + 1
        .Row = .Rows - 1


        For icol = 0 To rsDocs.Fields.Count - 1
            .Col = icol
            .Text = rsDocs(icol) & ""
        Next
        rsDocs.MoveNext
    Wend
'emp_id  emp_na  address telephone   dob sex department  in_cop  in_num  bsal

    .TextMatrix(0, 0) = "Employe ID"
    .TextMatrix(0, 1) = "Name"
    .TextMatrix(0, 2) = "Address"
    .TextMatrix(0, 3) = "Telephone"
    .TextMatrix(0, 4) = "DOB"
    .TextMatrix(0, 5) = "Sex"
    .TextMatrix(0, 6) = "Department"
    .TextMatrix(0, 7) = "Insurance Coper.."
    .TextMatrix(0, 8) = "Insureance Num"
    .TextMatrix(0, 9) = "Basic Sal"
        
    .FixedRows = 1
    .RowHeight(0) = .RowHeight(1) * 1.5
    .ColWidth(0) = .ColWidth(0) * 1.15
    .ColWidth(1) = .ColWidth(0) * 2.6
    .ColWidth(2) = .ColWidth(0) * 1.5
    .ColWidth(3) = .ColWidth(0) * 1.1
    .ColWidth(4) = .ColWidth(0) * 1.1
    .ColWidth(6) = .ColWidth(0) * 1.3
    .ColWidth(8) = .ColWidth(0) * 1.2
    .ColWidth(9) = .ColWidth(0) * 1.2
    .ColWidth(7) = .ColWidth(0) * 1.2
End With





rsDocs.Close
End If


End Sub

'##ModelId=3CBB305B038A
Private Sub MSFlexGrid1_DblClick()

Set con = New ADODB.Connection
Set rsDocs = New ADODB.Recordset
rsDocs.Open " Select * from Employee", cnPatients, adOpenDynamic, adLockOptimistic

rsDocs.MoveFirst

While Not rsDocs.EOF

If rsDocs.Fields(0) = MSFlexGrid1.Text Then
Call show_data
ElseIf rsDocs.Fields(1) = MSFlexGrid1.Text Then
Call show_data
ElseIf rsDocs.Fields(2) = MSFlexGrid1.Text Then
Call show_data
ElseIf rsDocs.Fields(3) = MSFlexGrid1.Text Then
Call show_data
ElseIf rsDocs.Fields(4) = MSFlexGrid1.Text Then
Call show_data
ElseIf rsDocs.Fields(5) = MSFlexGrid1.Text Then
Call show_data
ElseIf rsDocs.Fields(6) = MSFlexGrid1.Text Then
Call show_data
ElseIf rsDocs.Fields(7) = MSFlexGrid1.Text Then
Call show_data
ElseIf rsDocs.Fields(8) = MSFlexGrid1.Text Then
Call show_data
ElseIf rsDocs.Fields(9) = MSFlexGrid1.Text Then
Call show_data
End If
rsDocs.MoveNext
Wend


 

End Sub

'##ModelId=3CBB305B0393
Public Sub show_data()
With frm_add_employee
.txt_empid.Text = rsDocs.Fields(0)
.txt_emp_name.Text = rsDocs.Fields(1)
.txt_emp_address.Text = rsDocs.Fields(2)
.txt_emp_telephone.Text = rsDocs.Fields(3)
.DTPicker1.Value = rsDocs.Fields(4)
.cmb_emp_department.Text = rsDocs.Fields(6)
.txt_emp_insurecorp.Text = rsDocs.Fields(7)
.txt_emp_insurno.Text = rsDocs.Fields(8)
.txt_emp_bsal.Text = rsDocs.Fields(9)

If rsDocs.Fields(5) = "male" Then
.Option1(0).Value = True
.Option1(1).Value = False
ElseIf rsDocs.Fields(5) = "female" Then
.Option1(0).Value = False
.Option1(1).Value = True
End If
End With
End Sub
