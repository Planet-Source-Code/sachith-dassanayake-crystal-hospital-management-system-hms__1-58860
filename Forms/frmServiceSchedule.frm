VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmServiceSchedule 
   BackColor       =   &H00FF8080&
   Caption         =   "Hospital Service Schedule Details"
   ClientHeight    =   10560
   ClientLeft      =   1110
   ClientTop       =   450
   ClientWidth     =   10875
   Icon            =   "frmServiceSchedule.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10560
   ScaleWidth      =   10875
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "Schedule Details"
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
      Height          =   3855
      Left            =   1080
      TabIndex        =   23
      Top             =   1200
      Width           =   6375
      Begin VB.TextBox txtDay 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   2760
         Visible         =   0   'False
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker DTPin 
         Height          =   375
         Left            =   2160
         TabIndex        =   41
         Top             =   1560
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   45875203
         UpDown          =   -1  'True
         CurrentDate     =   38380
      End
      Begin MSComCtl2.DTPicker DTPout 
         Height          =   375
         Left            =   2160
         TabIndex        =   42
         Top             =   2160
         Visible         =   0   'False
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   45875203
         UpDown          =   -1  'True
         CurrentDate     =   38380
      End
      Begin VB.TextBox txtfields 
         DataField       =   "Service_Schedule_ID"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   32
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtfields 
         DataField       =   "Service_Ends"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   30
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox txtfields 
         DataField       =   "Service_Starts"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   29
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox txtfields 
         DataField       =   "Service_AvaiDate"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   2160
         TabIndex        =   28
         Top             =   2760
         Width           =   3375
      End
      Begin VB.TextBox txtfields 
         DataField       =   "Schedule_Notes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   27
         Top             =   3240
         Width           =   3375
      End
      Begin VB.ComboBox cmbDoctorID 
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
         ItemData        =   "frmServiceSchedule.frx":57E2
         Left            =   2160
         List            =   "frmServiceSchedule.frx":57E4
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   960
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CommandButton cmdShowAll 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5640
         TabIndex        =   24
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtfields 
         DataField       =   "Service_ID"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   31
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtTemp 
         DataField       =   "Service_ID"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   26
         Top             =   960
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Schedule ID:"
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
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   38
         Top             =   420
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Service ID:"
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
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   37
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Service Ends:"
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
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   36
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Service Starts:"
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
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   35
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Avaiable Days:"
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
         Height          =   285
         Index           =   4
         Left            =   240
         TabIndex        =   34
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Schedule Notes:"
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
         Height          =   285
         Index           =   5
         Left            =   240
         TabIndex        =   33
         Top             =   3360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Record Navigation"
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
      Height          =   1215
      Left            =   2040
      TabIndex        =   17
      Top             =   6600
      Width           =   7215
      Begin VB.CommandButton cmdFirst 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   360
         Picture         =   "frmServiceSchedule.frx":57E6
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   1065
         Picture         =   "frmServiceSchedule.frx":5CBC
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   5520
         Picture         =   "frmServiceSchedule.frx":619D
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdLast 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   6240
         Picture         =   "frmServiceSchedule.frx":6678
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   405
         Left            =   1800
         TabIndex        =   22
         Top             =   480
         Width           =   3480
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Record Operations"
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
      Height          =   1215
      Left            =   2160
      TabIndex        =   9
      Top             =   8040
      Width           =   6975
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
         Height          =   780
         Left            =   1560
         Picture         =   "frmServiceSchedule.frx":6B4D
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
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
         Height          =   780
         Left            =   5520
         Picture         =   "frmServiceSchedule.frx":7051
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   4200
         Picture         =   "frmServiceSchedule.frx":755D
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   2880
         Picture         =   "frmServiceSchedule.frx":7A03
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
         Height          =   780
         Left            =   1560
         Picture         =   "frmServiceSchedule.frx":7EBC
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         Height          =   780
         Left            =   240
         Picture         =   "frmServiceSchedule.frx":8361
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   240
         Picture         =   "frmServiceSchedule.frx":8810
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame frameDays 
      BackColor       =   &H00FF8080&
      Caption         =   "Available Days"
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
      Height          =   3855
      Left            =   7680
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CheckBox chkDay 
         BackColor       =   &H00FF8080&
         Caption         =   "Wednesday"
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
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   1800
         Width           =   1455
      End
      Begin VB.CheckBox chkDay 
         BackColor       =   &H00FF8080&
         Caption         =   "Saturday"
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
         Height          =   495
         Index           =   6
         Left            =   240
         TabIndex        =   7
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CheckBox chkDay 
         BackColor       =   &H00FF8080&
         Caption         =   "Friday"
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
         Height          =   495
         Index           =   5
         Left            =   240
         TabIndex        =   6
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CheckBox chkDay 
         BackColor       =   &H00FF8080&
         Caption         =   "Thursday"
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
         Height          =   495
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CheckBox chkDay 
         BackColor       =   &H00FF8080&
         Caption         =   "Tuesday"
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
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CheckBox chkDay 
         BackColor       =   &H00FF8080&
         Caption         =   "Monday"
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
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.CheckBox chkDay 
         BackColor       =   &H00FF8080&
         Caption         =   "Sunday"
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
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid grdDataGrid 
      Height          =   1305
      Left            =   120
      TabIndex        =   0
      Top             =   5280
      Width           =   10560
      _ExtentX        =   18627
      _ExtentY        =   2302
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
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
      FormHeightDT    =   11070
      FormWidthDT     =   10995
      FormScaleHeightDT=   10560
      FormScaleWidthDT=   10875
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "HOSPITAL SERVICES DETAILS"
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
      Left            =   2520
      TabIndex        =   39
      Top             =   360
      Width           =   6045
   End
End
Attribute VB_Name = "frmServiceSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Dim strDays As String

Private Sub chkDay_Click(Index As Integer)
Select Case (Index)

Case 0
If chkDay(0).Value = 1 Then
    strDays = strDays & "Sun,"
    'txtFields(4) = strDays
    txtDay = strDays
ElseIf chkDay(0).Value = 0 Then
 strDays = Replace(strDays, "Sun,", "")
 'txtFields(4).Text = strDays
 txtDay = strDays
End If

Case 1
If chkDay(1).Value = 1 Then
    strDays = strDays & "Mon,"
    'txtFields(4) = strDays
    txtDay = strDays
ElseIf chkDay(1).Value = 0 Then
 strDays = Replace(strDays, "Mon,", "")
 'txtFields(4).Text = strDays
 txtDay = strDays
End If


Case 2
If chkDay(2).Value = 1 Then
    strDays = strDays & "Tue,"
    'txtFields(4) = strDays
    txtDay = strDays
ElseIf chkDay(2).Value = 0 Then
 strDays = Replace(strDays, "Tue,", "")
 'txtFields(4).Text = strDays
 txtDay = strDays
End If


Case 3
If chkDay(3).Value = 1 Then
    strDays = strDays & "Wed,"
    'txtFields(4) = strDays
    txtDay = strDays
ElseIf chkDay(3).Value = 0 Then
 strDays = Replace(strDays, "Wed,", "")
 'txtFields(4).Text = strDays
 txtDay = strDays
End If


Case 4
If chkDay(4).Value = 1 Then
    strDays = strDays & "Thu,"
    'txtFields(4) = strDays
    txtDay = strDays
ElseIf chkDay(4).Value = 0 Then
 strDays = Replace(strDays, "Thu,", "")
 'txtFields(4).Text = strDays
 txtDay = strDays
End If


Case 5
If chkDay(5).Value = 1 Then
    strDays = strDays & "Fri,"
    'txtFields(4) = strDays
    txtDay = strDays
ElseIf chkDay(5).Value = 0 Then
 strDays = Replace(strDays, "Fri,", "")
 'txtFields(4).Text = strDays
 txtDay = strDays
End If


Case 6
If chkDay(6).Value = 1 Then
    strDays = strDays & "Sat,"
    'txtFields(4) = strDays
    txtDay = strDays
ElseIf chkDay(6).Value = 0 Then
 strDays = Replace(strDays, "Sat,", "")
 'txtFields(4).Text = strDays
 txtDay = strDays
End If


Case Else

End Select

End Sub



Private Sub cmdViewAll_Click()

End Sub

Private Sub Form_Activate()
Call DisableMenu
End Sub

Private Sub Form_Deactivate()
Call EnableMenu
End Sub

Private Sub Form_Load()
Me.WindowState = vbMaximized
  
  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "SHAPE {select Service_Schedule_ID,Service_ID,Service_Ends,Service_Starts,Service_AvaiDate,Schedule_Notes from Service_Schedule_Details Order by Service_Schedule_ID} AS ParentCMD APPEND ({select Service_Schedule_ID,Service_ID,Service_Starts,Service_Ends,Service_AvaiDate,Schedule_Notes from Service_Schedule_Details Order by Service_Starts } AS ChildCMD RELATE Service_ID TO Service_ID) AS ChildCMD", cnPatients, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
  
 Dim rsSched As Recordset
  Dim i As Integer
  Set rsSched = New ADODB.Recordset
  
    rsSched.Open "SELECT * FROM Services", cnPatients, adOpenKeyset, adLockPessimistic
    i = 0
    ' Add ID's to Combo Box
    
    While rsSched.EOF = False
        cmbDoctorID.AddItem rsSched(0)
        cmbDoctorID.ListIndex = i
        i = i + 1
        rsSched.MoveNext
    Wend
    rsSched.Close
  
  

  Set grdDataGrid.DataSource = adoPrimaryRS("ChildCMD").UnderlyingValue

  mbDataChanged = False
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call EnableMenu
  Screen.MousePointer = vbDefault
End Sub

Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
  lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  'On Error GoTo AddErr
  Dim SchedID As String
  Dim flag As Boolean
  Dim rsSched As Recordset
  Dim i As Integer
  Set rsSched = New ADODB.Recordset
  
   For i = 0 To 6
    chkDay(i).Value = 0
   Next i
   DTPin.Value = Format("00:00", "short time")
   DTPout.Value = Format("00:00", "short time")
   
        SchedID = Functions.UID(6, "SSchedID_")
        rsSched.Open "Select * from Service_Schedule_Details", cnPatients, adOpenKeyset, adLockPessimistic
              
            While rsSched.EOF = False
                If rsSched(0) = DocID Then
                   SchedID = Functions.UID(6, "SSchedID_")
                   rsSched.MoveFirst
                    flag = True
                Else
                    flag = False
                End If
                rsSched.MoveNext
            Wend
            
  
            rsSched.Close

            
  
   
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    
    cmbDoctorID.Visible = True
    'txtFields(1).Visible = False
    txtFields(1) = cmbDoctorID.Text
    txtTempDocID = cmbDoctorID.Text
    
    'Debug.Print txtFields(1)
    
    lblStatus.Caption = "Add record"
    frameDays.Visible = True
    txtFields(0).Text = SchedID
    For Each oText In Me.txtFields
        oText.Locked = False
    Next
    
    txtFields(0).Locked = True
    txtFields(1).Locked = True
    
    
    
    mbAddNewFlag = True
    SetButtons False
  End With

  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
   If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "Confirm Delete") = vbNo Then
    Exit Sub
  End If
  
  
  
  With adoPrimaryRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Set grdDataGrid.DataSource = Nothing
  adoPrimaryRS.Requery
  Set grdDataGrid.DataSource = adoPrimaryRS("ChildCMD").UnderlyingValue
  
  
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  Set grdDataGrid.DataSource = Nothing
  adoPrimaryRS.Requery
  Set grdDataGrid.DataSource = adoPrimaryRS("ChildCMD").UnderlyingValue
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()


  
    
    Dim oText As TextBox
  On Error GoTo EditErr
   Dim strDays As String
   Dim arrDays() As String
      
    cmbDoctorID = txtFields(1)
    
    strDays = txtFields(4)
    arrDays() = Split(strDays, ",")
    
    For i = 0 To 6
        chkDay(i).Value = 0
    Next
    
    For i = 0 To UBound(arrDays)
               
        If arrDays(i) = "Sun" Then
            chkDay(0).Value = 1
        End If
        If arrDays(i) = "Mon" Then
            chkDay(1).Value = 1
        End If
        If arrDays(i) = "Tue" Then
           chkDay(2).Value = 1
        End If
        If arrDays(i) = "Wed" Then
            chkDay(3).Value = 1
        End If
        If arrDays(i) = "Thu" Then
            chkDay(4).Value = 1
        End If
        If arrDays(i) = "Fri" Then
            chkDay(5).Value = 1
        End If
        If arrDays(i) = "Sat" Then
            chkDay(6).Value = 1
        End If
        
    Next i
    
    
    
    
    
    
    lblStatus.Caption = "Edit record"
    
    frameDays.Visible = True
    For Each oText In Me.txtFields
        oText.Locked = False
    Next
    txtFields(0).Locked = True
  
   DTPin.Value = txtFields(3)
  DTPout.Value = txtFields(2)
  cmbDoctorID.Visible = True
  cmbDoctorID.Text = txtFields(1).Text
  txtTempDocID = cmbDoctorID.Text
  
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  Dim oText As TextBox
  On Error Resume Next

  frameDays.Visible = False
    For Each oText In Me.txtFields
        oText.Locked = True
    Next
    txtFields(0).Locked = True
    txtFields(1).Visible = True
    cmbDoctorID.Visible = False
  
  
  
  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  
  End If
  
    mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
Dim oText As TextBox
  On Error GoTo UpdateErr
  
    txtFields(1) = cmbDoctorID.Text
    txtFields(3) = Format(DTPin, "Short time")
    txtFields(2) = Format(DTPout, "short time")
    txtFields(4) = txtDay
    
    
    If Right(txtFields(4), 1) = "," Then
        txtFields(4).Text = Left(txtFields(4), Len(txtFields(4)) - 1)
    End If
    
  adoPrimaryRS.UpdateBatch adAffectAll

  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'move to the new record
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False
  
  For Each oText In Me.txtFields
    oText.Locked = True
  Next
  
  txtFields(1).Visible = True
  cmbDoctorID.Visible = False
  
  
  frameDays.Visible = False
  
  
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  adoPrimaryRS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox Err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  adoPrimaryRS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox Err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError

  If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    adoPrimaryRS.MoveLast
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError

  If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveFirst
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox Err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdShowAll.Visible = Not bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
  txtFields(4).Locked = Not bVal
  txtFields(1).Locked = Not bVal
  'txtTempDocID.Visible = Not bVal
  
  
  DTPin.Visible = Not bVal
  DTPout.Visible = Not bVal
  txtDay.Visible = Not bVal
  
  
End Sub



