VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmBackUp 
   BackColor       =   &H00FF8080&
   Caption         =   "Backup Database"
   ClientHeight    =   6060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6645
   ForeColor       =   &H00800000&
   Icon            =   "frmBackUp.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6060
   ScaleWidth      =   6645
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   5280
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      AutoCenterForm  =   -1  'True
      FormHeightDT    =   6570
      FormWidthDT     =   6765
      FormScaleHeightDT=   6060
      FormScaleWidthDT=   6645
   End
   Begin VB.Frame frameCurrBackUp 
      BackColor       =   &H00FF8080&
      Caption         =   "Choose Path for BackUp"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   6375
      Begin VB.FileListBox File1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   480
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
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
         Left            =   4800
         Picture         =   "frmBackUp.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&BackUp"
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
         Left            =   4800
         Picture         =   "frmBackUp.frx":0DCE
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.DirListBox Dir1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1665
         Left            =   480
         TabIndex        =   1
         Top             =   960
         Width           =   3015
      End
      Begin VB.DriveListBox Drive1 
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
         Left            =   480
         TabIndex        =   0
         Top             =   480
         Width           =   3015
      End
   End
   Begin VB.Frame frameBackup 
      BackColor       =   &H00FF8080&
      Caption         =   "Last BackUp Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6375
      Begin VB.Label lblLastTime 
         BackColor       =   &H00FF8080&
         Caption         =   "Last BackUp Path"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   11
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label lblLastDate 
         BackColor       =   &H00FF8080&
         Caption         =   "Last BackUp Path"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label Time 
         BackColor       =   &H00FF8080&
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00FF8080&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblPath 
         BackColor       =   &H00FF8080&
         Caption         =   "Path"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblLastPath 
         BackColor       =   &H00FF8080&
         Caption         =   "Last BackUp Path"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1200
         TabIndex        =   5
         Top             =   480
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmBackUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Note : Microsoft Scripting Runtime Library is Referenced
'       For Making Object of File System Object


Dim Fsys As New FileSystemObject
Dim bckupFile As File

'Reading Previously Backup Details
Private Sub Form_Load()
    Call DisableMenu
    
    Dim lastPath As String
    Dim lastDate As String
    Dim lastTime As String
    
    'Read Registry for previous settings stored
    lastPath = GetSetting(App.Title, "Settings", "BackupPath")
    lastDate = GetSetting(App.Title, "Settings", "BackupDate")
    lastTime = GetSetting(App.Title, "Settings", "BackupTime")
    
    If lastPath = "" Then
        lblLastPath.Caption = "No Backup made previously"
        lblLastDate.Caption = " "
        lblLastTime.Caption = " "
    Else
        lblLastPath.Caption = lastPath
        lblLastDate.Caption = lastDate & "  (mm-dd-yy)"
        lblLastTime.Caption = lastTime
    End If
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Backup Cmd Btn
Private Sub cmdsave_click()
On Error Resume Next
    Dim dbname As String
    
    dbname = "HMS-" & Format$(Date, "mm-dd-yyyy") & ".mdb"
    cmdSave.Enabled = False
    Label1.Caption = "Please Wait, Backup in Progress..."
    Label1.BackColor = vbGreen
    Label1.ForeColor = vbYellow
    Dim destination As String
    Dim source As String
    Dim currDate, currTime As String
    currDate = Format$(Now, "mm - dd - yy")
    currTime = Format$(Now, "hh:mm:ss AM/PM")
    
    destination = File1.Path & "\" & dbname
    'destination = File1.Path & "\" & "HMS2.mdb"
    
    source = App.Path & "\database\HMS.mdb"
    
    'MsgBox "Source : " & source
    'MsgBox "Destination : " & destination
    Set bckupFile = Fsys.GetFile(finalpath)
    bckupFile.Attributes = Compressed
    Fsys.CopyFile source, destination, True
    'Saving Current Backup Details
    SaveSetting App.Title, "Settings", "BackupPath", destination
    SaveSetting App.Title, "Settings", "BackupDate", currDate
    SaveSetting App.Title, "Settings", "BackupTime", currTime
    
    
    cmdSave.Enabled = True
    
    'Dim strC As String
    'Set dbcon = New ADODB.Connection
    
    'strC = "provider=MSDataShape;Data provider=Microsoft.Jet.OLEDB.4.0;data source=destination;"
    
    'dbcon.CursorLocation = adUseClient
    'dbcon.Open (strC)
    
    'If Not dbcon.State = adStateOpen Then
        'MsgBox "Backup Process not successfull.", vbCritical, "Database Error"
    'Else
        'dbcon.Close
        MsgBox "BackUp Process Over", vbInformation, "Backup"
    'End If
    
    
       
    Unload Me
End Sub

Private Sub Drive1_Change()
    On Error Resume Next
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Dir1_Change()
    On Error Resume Next
    File1.Path = Dir1.Path
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call EnableMenu
End Sub
