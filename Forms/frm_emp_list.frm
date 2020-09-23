VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frm_emp_list 
   BackColor       =   &H00FF8080&
   Caption         =   "Employee List"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   8505
   Icon            =   "frm_emp_list.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   8505
   Begin VB.CommandButton cmd_doc_back 
      Appearance      =   0  'Flat
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
      Height          =   975
      Left            =   7200
      Picture         =   "frm_emp_list.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7680
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   11245
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16761024
      BackColorSel    =   16711680
      BackColorBkg    =   16744576
      GridColor       =   0
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
      FormHeightDT    =   9435
      FormWidthDT     =   8625
      FormScaleHeightDT=   8925
      FormScaleWidthDT=   8505
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
      Left            =   1800
      TabIndex        =   2
      Top             =   360
      Width           =   5220
   End
End
Attribute VB_Name = "frm_emp_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private rsDocs As New Recordset
'##ModelId=3CBB305B0324
Private strst As String
'##ModelId=3CBB305B032E
Private Sub Form_Load()


Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
'con.CursorLocation = adUseClient
'con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\pay\Pay.mdb;persist security info=false"
strst = "del"


Set rsDocs = New ADODB.Recordset


'create sql statement


rsDocs.Open "select ecode,name,dep,basic from slip  ", cnPatients, adOpenDynamic, adLockOptimistic



If rsDocs.EOF = False Then

With MSFlexGrid1
    .clear
    .Rows = 1
    .Cols = 4
   
  

    While Not rsDocs.EOF
        .Rows = .Rows + 1
        .Row = .Rows - 1


        For icol = 0 To MSFlexGrid1.Cols - 1
            .Col = icol
            .Text = rsDocs(icol) & ""
        Next
        rsDocs.MoveNext
    Wend
'    desig   Address basic   da  hra cca trans   gpf ins itax    ptax    gross   deduct  net
'ecode   name    desig   Address basic   da  hra cca trans   gpf ins itax    ptax    gross   deduct  net
    .TextMatrix(0, 0) = "Employe  Code"
    .TextMatrix(0, 1) = "Name"
    .TextMatrix(0, 2) = "Designation"
    .TextMatrix(0, 3) = "Basic Salary"
       
    .RowHeight(0) = .RowHeight(1) * 1.5
    .ColWidth(0) = .ColWidth(0) * 1.25
    .ColWidth(1) = .ColWidth(0) * 3
    .ColWidth(2) = .ColWidth(0) * 1.45
    .ColWidth(3) = .ColWidth(0) * 1.2
        
End With



rsDocs.Close

End If


End Sub
Private Sub cmd_doc_back_Click()
Unload Me
frm_add_salary_info.cmd_brows.SetFocus
End Sub




Private Sub MSFlexGrid1_Click()

rsDocs.Open "select * from slip ", cnPatients, adOpenDynamic, adLockOptimistic

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
rsDocs.Close
End Sub

Public Sub show_data()
With frm_add_salary_info
.Code.Text = rsDocs!ecode
.nm.Text = rsDocs!name
.desig.Text = rsDocs!dep
.add.Text = rsDocs!address
.bp.Text = rsDocs!basic
.da.Text = rsDocs!da
.hr.Text = rsDocs!hra
.ca.Text = rsDocs!cca
.ta.Text = rsDocs!trans
.pf.Text = rsDocs!gpf
.ins.Text = rsDocs!ins
.it.Text = rsDocs!itax
.pt.Text = rsDocs!ptax
.gpa.Text = rsDocs!gross
.deduct.Text = rsDocs!deduct
.Net.Text = rsDocs!Net
End With
End Sub
