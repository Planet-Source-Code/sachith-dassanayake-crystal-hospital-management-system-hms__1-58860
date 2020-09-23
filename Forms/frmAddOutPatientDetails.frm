VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmAddOutPatientDetails 
   BackColor       =   &H00FF8080&
   Caption         =   "Out Patient Details"
   ClientHeight    =   9930
   ClientLeft      =   4620
   ClientTop       =   4200
   ClientWidth     =   10200
   Icon            =   "frmAddOutPatientDetails.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9930
   ScaleWidth      =   10200
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF8080&
      Caption         =   "Patient Details"
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
      Height          =   7215
      Left            =   120
      TabIndex        =   38
      Top             =   960
      Width           =   7575
      Begin VB.TextBox txtFields 
         DataField       =   "Gender"
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
         Left            =   3000
         TabIndex        =   4
         Tag             =   "Chr"
         ToolTipText     =   "Gender"
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Patient_ID"
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
         Left            =   3000
         TabIndex        =   1
         ToolTipText     =   "Patient ID"
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtFields 
         DataField       =   "First_Name"
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
         Left            =   3000
         TabIndex        =   2
         Tag             =   "Chr"
         ToolTipText     =   "Patient First Name"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Last_Name"
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
         Left            =   3000
         TabIndex        =   3
         Tag             =   "Chr"
         ToolTipText     =   "Patient Last Name"
         Top             =   1845
         Width           =   2655
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Telephone"
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
         Left            =   3000
         TabIndex        =   5
         ToolTipText     =   "Patient Contact Number"
         Top             =   3000
         Width           =   2655
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Index           =   5
         Left            =   3000
         MultiLine       =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "Patient Address"
         Top             =   3480
         Width           =   2655
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Status"
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
         Index           =   6
         Left            =   3000
         TabIndex        =   7
         ToolTipText     =   "Patient Status"
         Top             =   4560
         Width           =   2655
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Notes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   7
         Left            =   3000
         MultiLine       =   -1  'True
         TabIndex        =   8
         ToolTipText     =   "Additional Notes"
         Top             =   5040
         Width           =   2655
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FF8080&
         Height          =   495
         Left            =   3000
         TabIndex        =   39
         Top             =   2280
         Visible         =   0   'False
         Width           =   2655
         Begin VB.OptionButton OptGender 
            BackColor       =   &H00FF8080&
            Caption         =   "Female"
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
            Index           =   1
            Left            =   1320
            TabIndex        =   41
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton OptGender 
            BackColor       =   &H00FF8080&
            Caption         =   "Male"
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
            Index           =   0
            Left            =   120
            TabIndex        =   40
            Top             =   120
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Patient ID:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   49
         Top             =   780
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "First Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   48
         Top             =   1380
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Last Name:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   47
         Top             =   1980
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Gender:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   46
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Telephone:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   45
         Top             =   3060
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   5
         Left            =   1080
         TabIndex        =   44
         Top             =   3645
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Status:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   6
         Left            =   1080
         TabIndex        =   43
         Top             =   4620
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FF8080&
         Caption         =   "Notes:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   7
         Left            =   1080
         TabIndex        =   42
         Top             =   5220
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000018&
      DragMode        =   1  'Automatic
      DrawWidth       =   12
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   960
      Negotiate       =   -1  'True
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      ScaleHeight     =   2715
      ScaleWidth      =   6195
      TabIndex        =   20
      Top             =   2880
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton cmdFCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   780
         Left            =   3240
         Picture         =   "frmAddOutPatientDetails.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton cmdFFind 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Search"
         Height          =   780
         Left            =   1800
         Picture         =   "frmAddOutPatientDetails.frx":134E
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000000FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   325
         Left            =   5880
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   15
         Width           =   325
      End
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   25
         Text            =   "Enter your Text Here"
         Top             =   1320
         Width           =   4575
      End
      Begin VB.OptionButton optSearch 
         BackColor       =   &H80000018&
         Caption         =   "Telephone"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   4800
         TabIndex        =   24
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optSearch 
         BackColor       =   &H80000018&
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3360
         TabIndex        =   23
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton optSearch 
         BackColor       =   &H80000018&
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   22
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton optSearch 
         BackColor       =   &H80000018&
         Caption         =   "Patient ID"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000001&
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   80
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000001&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   6255
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FF8080&
      Caption         =   "Add Appointments"
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
      Left            =   7080
      TabIndex        =   27
      Top             =   8400
      Width           =   2895
      Begin VB.OptionButton optApp 
         BackColor       =   &H00FF8080&
         Caption         =   "Doctor"
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
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optApp 
         BackColor       =   &H00FF8080&
         Caption         =   "Hospital Service"
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
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton cmdAddtoAppointment 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2040
         Picture         =   "frmAddOutPatientDetails.frx":17D1
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
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
      Left            =   120
      TabIndex        =   0
      Top             =   8400
      Width           =   6735
      Begin VB.CommandButton cmdFirst 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   240
         Picture         =   "frmAddOutPatientDetails.frx":1C4B
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdPrevious 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   960
         Picture         =   "frmAddOutPatientDetails.frx":2121
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   5160
         Picture         =   "frmAddOutPatientDetails.frx":2602
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.CommandButton cmdLast 
         BackColor       =   &H00FFFFFF&
         Height          =   540
         Left            =   5880
         Picture         =   "frmAddOutPatientDetails.frx":2ADD
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1680
         TabIndex        =   11
         Top             =   440
         Width           =   3360
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "Controls"
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
      Height          =   7215
      Left            =   7920
      TabIndex        =   32
      Top             =   960
      Width           =   2055
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
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
         Left            =   480
         Picture         =   "frmAddOutPatientDetails.frx":2FB2
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
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
         Left            =   480
         Picture         =   "frmAddOutPatientDetails.frx":3465
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1680
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
         Left            =   480
         Picture         =   "frmAddOutPatientDetails.frx":393C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2880
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
         Left            =   480
         Picture         =   "frmAddOutPatientDetails.frx":3DFF
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3960
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
         Left            =   480
         Picture         =   "frmAddOutPatientDetails.frx":42A5
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   6120
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Update"
         Height          =   780
         Left            =   480
         Picture         =   "frmAddOutPatientDetails.frx":47B1
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   780
         Left            =   480
         Picture         =   "frmAddOutPatientDetails.frx":4C5F
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1680
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdViewAll 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&View All"
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
         Left            =   480
         Picture         =   "frmAddOutPatientDetails.frx":5163
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   5040
         Width           =   1095
      End
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   480
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      FormHeightDT    =   10440
      FormWidthDT     =   10320
      FormScaleHeightDT=   9930
      FormScaleWidthDT=   10200
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "OUT PATIENTS"
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
      Left            =   3840
      TabIndex        =   35
      Top             =   240
      Width           =   2970
   End
End
Attribute VB_Name = "frmAddOutPatientDetails"
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

Private Sub cmdAddtoAppointment_Click()
Dim PatID As String

PatID = txtFields(0)
Unload Me

If optApp(0) = True Then
frmAddDocAppointments.cmbPatientID = PatID
frmAddDocAppointments.Show
ElseIf optApp(1) = True Then
frmAddSerAppointments.cmbPatientID = PatID
frmAddSerAppointments.Show
End If

Me.Hide


End Sub





Private Sub cmdFCancel_Click()
Picture1.Visible = False
Frame3.Enabled = True
End Sub

Private Sub cmdFFind_Click()
Dim strText As String
Dim SQL As String
'strText = InputBox("Please Enter The patient ID", "Search Patient", "OPID_")

strText = txtSearch
If optSearch(0) = True Then
    SearchFor = "Patient_ID"
ElseIf optSearch(1) = True Then
    SearchFor = "First_Name"
ElseIf optSearch(2) = True Then
    SearchFor = "Last_Name"
ElseIf optSearch(3) = True Then
    SearchFor = "Telephone"
End If

varBookMark = adoPrimaryRS.Bookmark
adoPrimaryRS.MoveFirst

SQL = SearchFor & "=" & "'" & strText & "'"

adoPrimaryRS.Find SQL

If (adoPrimaryRS.BOF = True) Or (adoPrimaryRS.EOF = True) Then
   MsgBox "Record not found"
   adoPrimaryRS.Bookmark = varBookMark
End If





End Sub

Private Sub cmdSearch_Click()
Picture1.Visible = True
Frame3.Enabled = False
End Sub

Private Sub cmdViewAll_Click()
frmDisplayOutPatient.Show
End Sub

Private Sub Command1_Click()
Picture1.Visible = False
Frame3.Enabled = True
End Sub

Private Sub Form_Activate()
Call Functions.DisableMenu
End Sub

Private Sub Form_Load()


Me.WindowState = vbMaximized


Call Functions.DisableMenu



  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select Patient_ID,First_Name,Last_Name,Gender,Telephone,Address,Status,Notes from Patient_Details ", cnPatients, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
    oText.Locked = True
  Next
  
  
  
  
  
  mbDataChanged = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 22 Then KeyAscii = 0: Exit Sub
    KeyAscii = DataEntryValidation(KeyAscii, ActiveControl.Tag)
End Sub


Private Sub Form_Resize()
  On Error Resume Next
  'lblStatus.Width = Me.Width - 1500
  'cmdNext.Left = lblStatus.Width + 700
  'cmdLast.Left = cmdNext.Left + 340
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
  Screen.MousePointer = vbDefault
  Call Functions.EnableMenu
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

  If bCancel Then
    adStatus = adStatusCancel
   
  End If
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
    
    Dim rsAddPatient As New Recordset
    Dim PID As String
    Set rsAddPatient = New ADODB.Recordset
  
    PID = Functions.UID(6, "OPID_")
    rsAddPatient.Open " Select * from Patient_Details", cnPatients, adOpenKeyset, adLockPessimistic
    While rsAddPatient.EOF = False
        If rsAddPatient(0) = PID Then
            ID = True
            PID = Functions.UID(6, "OPID_")
            rsAddPatient.MoveFirst
        Else
            ID = False
        End If
    rsAddPatient.MoveNext
    Wend
    rsAddPatient.Close
    
  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    oText.Locked = False
  Next
    
  txtFields(0).Locked = True
  
  
  
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    txtFields(0).Text = PID
    
    lblStatus.Caption = "Add record"
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
  Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  adoPrimaryRS.Requery
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr
  
  Dim oText As TextBox
  For Each oText In Me.txtFields
    oText.Locked = False
  Next
    
  txtFields(0).Locked = True
  

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next
  Dim oText As TextBox
  For Each oText In Me.txtFields
    oText.Locked = True
  Next
  

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  
  Picture1.Visible = False
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False

End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr
  Dim fl As Integer
  Dim oText As TextBox
  For Each oText In Me.txtFields
    oText.Locked = True
  Next
   fl = 0
   If txtFields(1) = "" Then
        MsgBox "Please Enter Patient First Name"
        txtFields(1).SetFocus
        fl = 1
    ElseIf txtFields(2) = "" Then
        MsgBox "Please Enter Patient Last Name"
        fl = 1
        txtFields(2).SetFocus
    ElseIf txtFields(5) = "" Then
        MsgBox "Please Enter Address"
        fl = 1
        txtFields(5).SetFocus
    ElseIf txtFields(4) = "" Then
        MsgBox "Please Enter Contact Number"
        fl = 1
        txtFields(4).SetFocus
    ElseIf txtFields(6) = "" Then
        MsgBox "Please Enter Patient Status"
        fl = 1
        txtFields(6).SetFocus
    ElseIf txtFields(0) = "" Then
        MsgBox "Please Enter Patient ID"
        fl = 1
    End If
    
    If OptGender(0).Value = True Then
        txtFields(3) = "Male"
    ElseIf OptGender(1).Value = True Then
        txtFields(3) = "FeMale"
    End If
  If fl <> 1 Then
    adoPrimaryRS.UpdateBatch adAffectAll
   
  
  If mbAddNewFlag Then
        adoPrimaryRS.MoveLast              'move to the new record
  End If
    mbEditFlag = False
    mbAddNewFlag = False
    SetButtons True
    Picture1.Visible = False
    mbDataChanged = False
  Exit Sub
  ElseIf fl = 1 Then
    For Each oText In Me.txtFields
        oText.Locked = False
    Next
        txtFields(0).Locked = True
    Exit Sub
    

  End If
 
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
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmdDelete.Visible = bVal
  cmdClose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdNext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
  Frame1.Visible = Not bVal
  txtFields(3).Visible = bVal
  cmdAddtoAppointment.Visible = bVal
  Picture1.Visible = bVal
  'cmdSearch.Visible = bVal
  Frame5.Visible = bVal
  
  cmdViewAll.Visible = bVal
  
End Sub

Private Sub OptGender_Click(Index As Integer)
Select Case (Index)
        Case "0" ' Male
        txtFields(3) = "Male"
        Case "1" 'Female
        txtFields(3) = "FeMale"

        Case Else 'None
            
End Select

End Sub



