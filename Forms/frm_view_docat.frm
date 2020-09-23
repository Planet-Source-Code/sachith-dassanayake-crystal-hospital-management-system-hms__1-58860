VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frm_view_docat 
   BackColor       =   &H00FF8080&
   Caption         =   "Avilable Categories"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   4650
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_view_docat.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   4650
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
      Left            =   3120
      Picture         =   "frm_view_docat.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5741
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16761024
      BackColorFixed  =   14737632
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
      FormHeightDT    =   6330
      FormWidthDT     =   4770
      FormScaleHeightDT=   5820
      FormScaleWidthDT=   4650
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "VIEW CATEOGRIES"
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
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   3810
   End
End
Attribute VB_Name = "frm_view_docat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3CBB305B02FC"
Private strok As String

Private Sub cmdClose_Click()
Unload Me
End Sub

'##ModelId=3CBB305B02FD
Private Sub Form_Load()

Dim rsDocs As New Recordset

Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
con.CursorLocation = adUseClient
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\pay\Pay.mdb;persist security info=false"
strok = "ok"


Set rsDocs = New ADODB.Recordset


'create sql statement


rsDocs.Open "select * from Doctor_special where recstate='" & strok & "' ", con, adOpenDynamic, adLockOptimistic



If rsDocs.EOF = False Then

With MSFlexGrid1
    .clear
    .Rows = 1
    .Cols = rsDocs.Fields.Count - 1
   
  

    While Not rsDocs.EOF
        .Rows = .Rows + 1
        .Row = .Rows - 1


        For icol = 0 To rsDocs.Fields.Count - 2
            .Col = icol
            .Text = rsDocs(icol) & ""
        Next
        rsDocs.MoveNext
    Wend
    
    
    .TextMatrix(0, 0) = "Department ID"
    .TextMatrix(0, 1) = "Department"
    

    .FixedRows = 1
    .RowHeight(0) = .RowHeight(1) * 1.5
    .ColWidth(0) = .ColWidth(0) * 1.75
    .ColWidth(1) = .ColWidth(0) * 2

End With



rsDocs.Close

End If


End Sub


