VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frm_salary_view 
   BackColor       =   &H00FF8080&
   Caption         =   "Salary Table View"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   13485
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_salary_view.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7890
   ScaleWidth      =   13485
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
      Left            =   12000
      Picture         =   "frm_salary_view.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   9340
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16761024
      BackColorBkg    =   16744576
      GridColor       =   0
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
      FormHeightDT    =   8400
      FormWidthDT     =   13605
      FormScaleHeightDT=   7890
      FormScaleWidthDT=   13485
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "VIEW SALARY DETAILS"
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
      Left            =   4560
      TabIndex        =   2
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "frm_salary_view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3CBB305B031A"
'##ModelId=3CBB305B0324
Private strst As String

Private Sub cmdClose_Click()
Unload Me
End Sub

'##ModelId=3CBB305B032E
Private Sub Form_Load()
Dim rsDocs As New Recordset

Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
'con.CursorLocation = adUseClient
'con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\pay\Pay.mdb;persist security info=false"
strst = "del"


Set rsDocs = New ADODB.Recordset


'create sql statement


rsDocs.Open "select * from slip  ", cnPatients, adOpenDynamic, adLockOptimistic



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
'    desig   Address basic   da  hra cca trans   gpf ins itax    ptax    gross   deduct  net
'ecode   name    desig   Address basic   da  hra cca trans   gpf ins itax    ptax    gross   deduct  net
    .TextMatrix(0, 0) = "Employe  Code"
    .TextMatrix(0, 1) = "Name"
    .TextMatrix(0, 2) = "Designation"
    .TextMatrix(0, 3) = "Address"
    .TextMatrix(0, 4) = "Basic Salary"
    .TextMatrix(0, 5) = "DA"
    .TextMatrix(0, 6) = "HRA"
    .TextMatrix(0, 7) = "CCA"
    .TextMatrix(0, 8) = "Trans"
    .TextMatrix(0, 9) = "GPF"
    .TextMatrix(0, 10) = "INS"
    .TextMatrix(0, 11) = "ITAX"
    .TextMatrix(0, 12) = "PTAX"
    .TextMatrix(0, 13) = "Gross"
    .TextMatrix(0, 14) = "Deduct"
    .TextMatrix(0, 15) = "Net Salary"
    
    .FixedRows = 1
    .RowHeight(0) = .RowHeight(1) * 1.5
    .ColWidth(1) = .ColWidth(0) * 2.5
    .ColWidth(2) = .ColWidth(0) * 1.75
    .ColWidth(3) = .ColWidth(0) * 1.5
    .ColWidth(4) = .ColWidth(0) * 1.5
End With



rsDocs.Close

End If


End Sub
