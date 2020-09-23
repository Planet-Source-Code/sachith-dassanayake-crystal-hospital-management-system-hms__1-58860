VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmbedwardreport 
   BackColor       =   &H00FF8080&
   Caption         =   "Reports"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11505
   Icon            =   "frmbedwardreport.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4275
   ScaleWidth      =   11505
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   1095
      Left            =   9720
      Picture         =   "frmbedwardreport.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Bed - Ward Reports"
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
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   11295
      Begin VB.CommandButton cmd_ward 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total &Ward Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1680
         Picture         =   "frmbedwardreport.frx":0DCE
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FF8080&
         Caption         =   "Custom Reports"
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
         Height          =   1455
         Left            =   3240
         TabIndex        =   4
         Top             =   240
         Width           =   7815
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FF8080&
            Caption         =   "Occupied Beds"
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
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FF8080&
            Caption         =   "Vacant Beds"
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
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   720
            Width           =   1575
         End
         Begin VB.CommandButton cmd_bed_custom 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "G&enerate"
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
            Left            =   6360
            Picture         =   "frmbedwardreport.frx":1379
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox cmb_bed_room 
            Appearance      =   0  'Flat
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
            Left            =   4080
            TabIndex        =   6
            Top             =   840
            Width           =   2055
         End
         Begin VB.ComboBox cmb_bed_ward 
            Appearance      =   0  'Flat
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
            Left            =   4080
            TabIndex        =   5
            Top             =   360
            Width           =   2055
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FF8080&
            Caption         =   "Ward Name"
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
            Left            =   2160
            TabIndex        =   11
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FF8080&
            Caption         =   "Room Type"
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
            Left            =   2160
            TabIndex        =   10
            Top             =   840
            Width           =   1215
         End
      End
      Begin VB.CommandButton cmd_gardian_report 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total &Guardian Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1680
         Picture         =   "frmbedwardreport.frx":14B1
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmd_inpa_report 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total &In-Patient Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   360
         Picture         =   "frmbedwardreport.frx":193A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton cmd_bed 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total &Bed Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   360
         Picture         =   "frmbedwardreport.frx":1DC3
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin Crystal.CrystalReport Report 
      Left            =   2520
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      FormHeightDT    =   4785
      FormWidthDT     =   11625
      FormScaleHeightDT=   4275
      FormScaleWidthDT=   11505
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "BED - WARD REPORTS"
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
      Left            =   3600
      TabIndex        =   13
      Top             =   240
      Width           =   4500
   End
End
Attribute VB_Name = "frmbedwardreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsWard As New ADODB.Recordset
Private rsRoom As New ADODB.Recordset
Private flag1 As String
Private flag2 As String
Private flag3 As String
Private flag4 As String
Private checkflag As Boolean



Private Sub cmb_bed_room_Click()
cmb_bed_ward.Text = ""
flag4 = cmb_bed_room.Text
checkflag = True
Debug.Print flag4, checkflag
End Sub

Private Sub cmb_bed_ward_Click()
cmb_bed_room.Text = ""
flag4 = cmb_bed_ward.Text
checkflag = False
Debug.Print flag4, checkflag
End Sub

Private Sub cmd_bed_Click()
Report.ReportFileName = App.Path & "\Reports\employee\BedDetail.rpt"
Report.ReplaceSelectionFormula ("")
Report.WindowState = crptMaximized
Report.Action = 1
End Sub

Private Sub cmd_bed_custom_Click()
If flag3 = "" And flag4 = "" Then
MsgBox "Plese Select options"
ElseIf flag3 = "Occu" And flag4 = "" Then

Report.ReportFileName = App.Path & "\Reports\employee\BedDetail.rpt"
Report.ReplaceSelectionFormula ("{Bed_Details.Available} = False")
Report.WindowState = crptMaximized
Report.Action = 1

ElseIf flag3 = "vac" And flag4 = "" Then

Report.ReportFileName = App.Path & "\Reports\employee\BedDetail.rpt"
Report.ReplaceSelectionFormula ("{Bed_Details.Available} = True")
Report.WindowState = crptMaximized
Report.Action = 1
'------------------------------------------------------------------------
ElseIf flag3 = "Occu" And checkflag = True Then

Report.ReportFileName = App.Path & "\Reports\employee\BedDetailcustom-1.rpt"
Report.ReplaceSelectionFormula ("{Bed_Details.Available} = False and {Bed_Details.Room_Ward_ID} = {Room_Details.Room_ID} and {Room_Details.Room_Type} = '" & cmb_bed_room.Text & "'")
Report.WindowState = crptMaximized
Report.Action = 1

ElseIf flag3 = "Occu" And checkflag = False Then

Report.ReportFileName = App.Path & "\Reports\employee\BedDetailcustom-1.rpt"
Report.ReplaceSelectionFormula ("{Bed_Details.Available} = False and {Bed_Details.Room_Ward_ID}  = {Ward_Details.Ward_ID} and {Ward_Details.Ward_Name} = '" & cmb_bed_ward.Text & "'")
Report.WindowState = crptMaximized
Report.Action = 1

'-----------------------------------------------------------------------
ElseIf flag3 = "vac" And checkflag = True Then

Report.ReportFileName = App.Path & "\Reports\employee\BedDetailcustom-1.rpt"
Report.ReplaceSelectionFormula ("{Bed_Details.Available} = True and {Bed_Details.Room_Ward_ID} = {Room_Details.Room_ID} and {Room_Details.Room_Type} = '" & cmb_bed_room.Text & "'")
Report.WindowState = crptMaximized
Report.Action = 1

ElseIf flag3 = "vac" And checkflag = False Then

Report.ReportFileName = App.Path & "\Reports\employee\BedDetailcustom-1.rpt"
Report.ReplaceSelectionFormula ("{Bed_Details.Available} = True and {Bed_Details.Room_Ward_ID}  = {Ward_Details.Ward_ID} and {Ward_Details.Ward_Name} = '" & cmb_bed_ward.Text & "'")
Report.WindowState = crptMaximized
Report.Action = 1

'{Bed_Details.Available} = False and{Bed_Details.Room_Ward_ID} = {Ward_Details.Ward_ID} and {Ward_Details.Ward_Name} = "Aids Ward"


End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Option3_Click(Index As Integer)

If Option3(0).Value = True Then
flag3 = "Occu"
ElseIf Option3(1).Value = True Then
flag3 = "vac"
End If

End Sub


Private Sub cmd_gardian_report_Click()
Report.ReportFileName = App.Path & "\Reports\employee\Gardian.rpt"
Report.DiscardSavedData = True
Report.WindowState = crptMaximized
Report.Action = 1
End Sub

Private Sub cmd_inpa_report_Click()
Report.ReportFileName = App.Path & "\Reports\employee\inpatient.rpt"
Report.DiscardSavedData = True
Report.WindowState = crptMaximized
Report.Action = 1
End Sub

Private Sub cmd_ward_Click()
Report.ReportFileName = App.Path & "\Reports\employee\WardDetails.rpt"
Report.DiscardSavedData = True
Report.ReplaceSelectionFormula ("")
Report.WindowState = crptMaximized
Report.Action = 1
End Sub

Private Sub Form_Load()

rsWard.Open "select Ward_Name from Ward_Details  ", cnPatients, adOpenDynamic, adLockOptimistic


rsWard.MoveFirst
While rsWard.EOF = False
cmb_bed_ward.AddItem rsWard!Ward_Name
rsWard.MoveNext
Wend
'Ward_ID Ward_Name   Ward_Rate   Ward_Desc
rsWard.Close


rsRoom.Open "select * from Room_Types  ", cnPatients, adOpenDynamic, adLockOptimistic

rsRoom.MoveFirst
While rsRoom.EOF = False
cmb_bed_room.AddItem rsRoom!Room_Type
rsRoom.MoveNext
Wend
rsRoom.Close


End Sub
