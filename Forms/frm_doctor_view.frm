VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frm_doctor_view 
   BackColor       =   &H00FF8080&
   Caption         =   "Doctor Table View"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   6585
   ClientWidth     =   14880
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_doctor_view.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   14880
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
      Left            =   13440
      Picture         =   "frm_doctor_view.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   14715
      _ExtentX        =   25956
      _ExtentY        =   8070
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16761024
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
      FormHeightDT    =   7890
      FormWidthDT     =   15000
      FormScaleHeightDT=   7380
      FormScaleWidthDT=   14880
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "VIEW DOCTORS DETAILS"
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
      Left            =   5280
      TabIndex        =   2
      Top             =   240
      Width           =   5025
   End
End
Attribute VB_Name = "frm_doctor_view"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3CBB305B034C"
Private strst As String

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim rsDocs As New Recordset

Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
con.CursorLocation = adUseClient
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\pay\Pay.mdb;persist security info=false"
strst = "del"


Set rsDocs = New ADODB.Recordset


'create sql statement


rsDocs.Open "select * from doctor_detail", con, adOpenDynamic, adLockOptimistic



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
'
'doc_id  doc_name    address telephone   mobile  dob sex speciality  in_na   in_num  bsal

    .TextMatrix(0, 0) = "Doctor ID"
    .TextMatrix(0, 1) = "Name"
    .TextMatrix(0, 2) = "Address"
    .TextMatrix(0, 3) = "Telephone"
    .TextMatrix(0, 4) = "Mobile"
    .TextMatrix(0, 5) = "DOB"
    .TextMatrix(0, 6) = "Sex"
    .TextMatrix(0, 7) = "Field"
    .TextMatrix(0, 8) = "Insureance Cope.."
    .TextMatrix(0, 9) = "Insureance Num"
    .TextMatrix(0, 10) = "Status"
    
    .FixedRows = 1
    .RowHeight(0) = .RowHeight(1) * 1.5
    .ColWidth(0) = .ColWidth(0) * 1.15
    .ColWidth(1) = .ColWidth(0) * 2
    .ColWidth(2) = .ColWidth(0) * 1.4
    .ColWidth(3) = .ColWidth(0) * 1.1
    .ColWidth(4) = .ColWidth(0) * 1.1
    .ColWidth(5) = .ColWidth(0) * 1.2
    .ColWidth(5) = .ColWidth(0) * 1.2
    .ColWidth(8) = .ColWidth(0) * 1.2
    .ColWidth(9) = .ColWidth(0) * 1.2
    .ColWidth(7) = .ColWidth(0) * 1.2
End With



rsDocs.Close

End If


End Sub





Private Sub MSFlexGrid1_Click()
rsDocs.Close
rsDocs.Open "select * from doctor_detail ", con, adOpenDynamic, adLockOptimistic

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
With frm_add_doctor
.txt_docid.Text = rsDocs.Fields(0)
.txt_doc_name.Text = rsDocs.Fields(1)
.txt_doc_address.Text = rsDocs.Fields(2)
.txt_Telephone.Text = rsDocs.Fields(3)
.txt_mobile.Text = rsDocs.Fields(4)
.DTPicker1.Value = rsDocs.Fields(5)
.cmb_doc_spec.Text = rsDocs.Fields(7)
.txt_doc_insurecorp.Text = rsDocs(8)
.txt_doc_insurno.Text = rsDocs.Fields(9)
.cmb_status.Text = rsDocs.Fields(10)

If rsDocs.Fields(6) = "male" Then
.Option1(0).Value = True
.Option1(1).Value = False
ElseIf rsDocs.Fields(6) = "female" Then
.Option1(0).Value = False
.Option1(1).Value = True
End If
End With
End Sub
