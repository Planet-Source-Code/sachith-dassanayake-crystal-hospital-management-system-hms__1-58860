VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmRoomDetails 
   BackColor       =   &H00FF8080&
   Caption         =   "View Room Details"
   ClientHeight    =   6045
   ClientLeft      =   3660
   ClientTop       =   3525
   ClientWidth     =   6855
   Icon            =   "frmRoomDetails.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   6855
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
      Height          =   885
      Left            =   5400
      Picture         =   "frmRoomDetails.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6165
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
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
      FormHeightDT    =   6555
      FormWidthDT     =   6975
      FormScaleHeightDT=   6045
      FormScaleWidthDT=   6855
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "VIEW ROOM DETAILS"
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
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   4320
   End
End
Attribute VB_Name = "frmRoomDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()
Call Functions.DisableMenu
Dim rsRoom As New Recordset

Dim SQL As String

Set rsRoom = New ADODB.Recordset


'create sql statement

SQL = "SELECT * FROM Room_Details"

'Set rsDocs.ActiveConnection = cnPatients

rsRoom.Open SQL, cnPatients, adOpenStatic, adLockPessimistic

If rsRoom.EOF = False Then

With MSFlexGrid1
    .clear
    .Rows = 1
    .Cols = rsRoom.Fields.Count
   
  

    While Not rsRoom.EOF
        .Rows = .Rows + 1
        .Row = .Rows - 1


        For icol = 0 To rsRoom.Fields.Count - 1
            .Col = icol
            .Text = rsRoom(icol) & ""
        Next
        rsRoom.MoveNext
    Wend
    
    
    .TextMatrix(0, 0) = "Room ID"
    .TextMatrix(0, 1) = "Room Type"
    .TextMatrix(0, 2) = "Description"
    
    .FixedRows = 1
    .RowHeight(0) = .RowHeight(1) * 1.5
     Functions.SizeColumns MSFlexGrid1, Me
End With



rsRoom.Close

End If



End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Functions.EnableMenu
End Sub
