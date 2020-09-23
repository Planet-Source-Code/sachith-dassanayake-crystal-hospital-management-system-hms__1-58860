VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmWardDetails 
   BackColor       =   &H00FF8080&
   Caption         =   "View Ward Details"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9405
   Icon            =   "frmWardDetails.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   9405
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
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
      Height          =   1005
      Left            =   7560
      Picture         =   "frmWardDetails.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3135
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5530
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
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
      Resolution      =   2
      ScreenHeight    =   768
      ScreenWidth     =   1024
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      FormHeightDT    =   6300
      FormWidthDT     =   9525
      FormScaleHeightDT=   5790
      FormScaleWidthDT=   9405
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "VIEW WARD DETAILS"
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
      Left            =   2760
      TabIndex        =   2
      Top             =   240
      Width           =   4395
   End
End
Attribute VB_Name = "frmWardDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()
Call Functions.DisableMenu


Dim rsWard As New Recordset

Dim SQL As String

Set rsWard = New ADODB.Recordset


'create sql statement

SQL = "SELECT * FROM Ward_Details"

'Set rsDocs.ActiveConnection = cnPatients

rsWard.Open SQL, cnPatients, ad, adLockPessimistic

If rsWard.EOF = False Then

With MSFlexGrid1
    .clear
    .Rows = 1
    .Cols = rsWard.Fields.Count
   
  

    While Not rsWard.EOF
        .Rows = .Rows + 1
        .Row = .Rows - 1


        For icol = 0 To rsWard.Fields.Count - 1
            .Col = icol
            .Text = rsWard(icol) & ""
        Next
        rsWard.MoveNext
    Wend
    
    
    .TextMatrix(0, 0) = "Ward ID"
    .TextMatrix(0, 1) = "Ward Type"
    .TextMatrix(0, 2) = "Ward Rate"
    .TextMatrix(0, 3) = "Description"
    
    .FixedRows = 1
    .RowHeight(0) = .RowHeight(1) * 1.5
     Functions.SizeColumns MSFlexGrid1, Me
End With



rsWard.Close

End If



End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Functions.EnableMenu
End Sub
