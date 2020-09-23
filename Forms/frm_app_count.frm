VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frm_app_count 
   Appearance      =   0  'Flat
   BackColor       =   &H00FF8080&
   Caption         =   "Doctor Salary Calculation"
   ClientHeight    =   10200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11835
   Icon            =   "frm_app_count.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10200
   ScaleWidth      =   11835
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF8080&
      Caption         =   "Month"
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
      Height          =   4095
      Left            =   6600
      TabIndex        =   47
      Top             =   840
      Width           =   2535
      Begin VB.ComboBox cmb_month 
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
         Height          =   2715
         ItemData        =   "frm_app_count.frx":57E2
         Left            =   240
         List            =   "frm_app_count.frx":580D
         Style           =   1  'Simple Combo
         TabIndex        =   48
         Top             =   600
         Width           =   1935
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "Doctor Details"
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
      Height          =   4095
      Left            =   360
      TabIndex        =   38
      Top             =   840
      Width           =   5895
      Begin VB.ComboBox cmb_doc_id 
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
         Left            =   3120
         TabIndex        =   42
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txt_doc_name 
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
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox txt_num_of_apps 
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
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txt_num_of_visits 
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
         Height          =   285
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor ID"
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
         Left            =   600
         TabIndex        =   46
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor Name"
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
         Left            =   600
         TabIndex        =   45
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Number Of Appointments"
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
         Left            =   600
         TabIndex        =   44
         Top             =   1320
         Width           =   2535
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Number Of Visits"
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
         Left            =   600
         TabIndex        =   43
         Top             =   1800
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Operations"
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
      Height          =   8895
      Left            =   9360
      TabIndex        =   30
      Top             =   840
      Width           =   1935
      Begin VB.CommandButton cmd_save 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
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
         Left            =   480
         Picture         =   "frm_app_count.frx":5873
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmd_clear 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Clear"
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
         Left            =   480
         Picture         =   "frm_app_count.frx":5D21
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   6360
         Width           =   1095
      End
      Begin VB.CommandButton cmd_view_doc 
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
         Height          =   975
         Left            =   480
         Picture         =   "frm_app_count.frx":61D2
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Search"
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
         Left            =   480
         Picture         =   "frm_app_count.frx":668D
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3960
         Width           =   1095
      End
      Begin VB.CommandButton cmd_back 
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
         Left            =   480
         Picture         =   "frm_app_count.frx":6B10
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   7560
         Width           =   1095
      End
      Begin VB.CommandButton cmd_del 
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
         Height          =   975
         Left            =   480
         Picture         =   "frm_app_count.frx":7014
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmd_modify 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Modify"
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
         Left            =   480
         Picture         =   "frm_app_count.frx":74CD
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2760
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Salary Details"
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
      Height          =   4455
      Left            =   360
      TabIndex        =   1
      Top             =   5280
      Width           =   8775
      Begin VB.TextBox Net 
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
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   15
         Tag             =   "Num"
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox deduct 
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
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   14
         Tag             =   "Num"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox gpa 
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
         Left            =   3000
         TabIndex        =   13
         Tag             =   "Num"
         Top             =   3840
         Width           =   1215
      End
      Begin VB.TextBox it 
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
         Height          =   285
         Left            =   7080
         MaxLength       =   5
         TabIndex        =   12
         Tag             =   "Num"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox ins 
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
         Height          =   285
         Left            =   7080
         MaxLength       =   5
         TabIndex        =   11
         Tag             =   "Num"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox pf 
         Appearance      =   0  'Flat
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
         Height          =   285
         Left            =   7080
         MaxLength       =   5
         TabIndex        =   10
         Tag             =   "Num"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox ta 
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
         Height          =   285
         Left            =   3000
         TabIndex        =   9
         Tag             =   "Num"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox ca 
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
         Height          =   285
         Left            =   3000
         TabIndex        =   8
         Tag             =   "Num"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox hr 
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
         Height          =   285
         Left            =   3000
         TabIndex        =   7
         Tag             =   "Num"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox da 
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
         Height          =   285
         Left            =   3000
         TabIndex        =   6
         Tag             =   "Num"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox bp 
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
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   5
         Tag             =   "Num"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox pt 
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
         Height          =   285
         Left            =   7080
         TabIndex        =   4
         Tag             =   "Num"
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox ch 
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   7080
         TabIndex        =   3
         Tag             =   "Num"
         Top             =   2880
         Width           =   1215
      End
      Begin VB.TextBox vc 
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   7080
         MaxLength       =   5
         TabIndex        =   2
         Tag             =   "Num"
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nett Pay"
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
         Height          =   195
         Left            =   600
         TabIndex        =   29
         Top             =   3360
         Width           =   825
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deductions"
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
         Height          =   195
         Left            =   600
         TabIndex        =   28
         Top             =   2880
         Width           =   1080
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Pay"
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
         Height          =   195
         Left            =   600
         TabIndex        =   27
         Top             =   3840
         Width           =   975
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P. Tax"
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
         Height          =   195
         Left            =   4800
         TabIndex        =   26
         Top             =   1920
         Width           =   600
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Income Tax"
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
         Height          =   195
         Left            =   4800
         TabIndex        =   25
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance"
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
         Height          =   195
         Left            =   4800
         TabIndex        =   24
         Top             =   960
         Width           =   990
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G.P.F"
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
         Height          =   195
         Left            =   4800
         TabIndex        =   23
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transport Allowance"
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
         Height          =   195
         Left            =   600
         TabIndex        =   22
         Top             =   2400
         Width           =   2010
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C.C.A."
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
         Height          =   195
         Left            =   600
         TabIndex        =   21
         Top             =   1920
         Width           =   555
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "H.R.A."
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
         Height          =   195
         Left            =   600
         TabIndex        =   20
         Top             =   1440
         Width           =   570
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D.A."
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
         Height          =   195
         Left            =   600
         TabIndex        =   19
         Top             =   960
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Basic Pay"
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
         Height          =   195
         Left            =   600
         TabIndex        =   18
         Top             =   480
         Width           =   930
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Visiting Charges"
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
         Height          =   195
         Left            =   4800
         TabIndex        =   17
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Channeling Charges"
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
         Height          =   195
         Left            =   4800
         TabIndex        =   16
         Top             =   2880
         Width           =   1935
      End
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
      FormHeightDT    =   10710
      FormWidthDT     =   11955
      FormScaleHeightDT=   10200
      FormScaleWidthDT=   11835
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DOCTOR SALARY DETAILS"
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
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frm_app_count"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public str_choice_1 As String
Public str_search_number As String
Public rs1 As New ADODB.Recordset
Public rs As New ADODB.Recordset
Public rscat As New ADODB.Recordset
Public rsdep As New ADODB.Recordset
Public rsDocs As New ADODB.Recordset
'Public cnPatients As New ADODB.Connection
Dim numapp As Integer
Dim numvis As Integer


Private Sub ca_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
    End If
End Sub

Private Sub cmb_doc_id_Change()
If Not cmb_doc_id = "" Then
Frame4.Visible = True
cmb_month.Visible = True
Else
Frame4.Visible = False
cmb_month.Visible = False
End If
End Sub

Private Sub cmb_doc_id_Click()
Dim na1 As String
Dim na2 As String

rsDocs.Open "select Distinct Doctor_ID, Doctor_Type,Doctor_FName,Doctor_LName,Doctor_Basic_sal from Doctor_Details where Doctor_ID= '" & cmb_doc_id.Text & "'", cnPatients, adOpenDynamic, adLockOptimistic
'Debug.Print rsDocs!Doctor_FName, rsDocs!Doctor_LName
na1 = Trim(rsDocs!Doctor_FName)
na2 = Trim(rsDocs!Doctor_LName)
txt_doc_name.Text = na1 + " " + na2

If Not cmb_doc_id = "" Then
Frame4.Visible = True
cmb_month.Visible = True

Else
Frame4.Visible = False
cmb_month.Visible = False
End If
bp = rsDocs!Doctor_Basic_sal
rsDocs.Close

rsDocs.Open "select * from Doctor_salary where decode= '" & cmb_doc_id.Text & "'", cnPatients, adOpenDynamic, adLockOptimistic
Debug.Print rsDocs.RecordCount
If rsDocs.RecordCount >= 1 Then
With frm_app_count
.bp = rsDocs!basic
.ca = rsDocs!cca
.ch = rsDocs!ch
.da = rsDocs!da
.deduct = rsDocs!deduct
.gpa = rsDocs!gross
.hr = rsDocs!hra
.ins = rsDocs!ins
.it = rsDocs!itax
.Net = rsDocs!Net
.pf = rsDocs!gpf
.pt = rsDocs!ptax
.ta = rsDocs!trans
.vc = rsDocs!vc
.txt_num_of_apps = ""
.txt_num_of_visits = ""
End With
Frame4.Visible = False
cmb_month.Visible = False

cmd_save.Enabled = False
Else
Call clear
cmd_save.Enabled = True
End If


rsDocs.Close
End Sub



Private Sub cmb_month_Click()
Dim date2 As Integer

Dim i As Integer
Dim j As Integer
i = 0
j = 0
Dim rsrec As Recordset
Set rsrec = New ADODB.Recordset
Dim rsvis As New ADODB.Recordset
date2 = 0
date2 = Val(cmb_month.ListIndex) + 1

'MsgBox Combo2.ListCount
'MsgBox date2


rsrec.Open "Select * from Doctor_Appointment where Month(Appointment_Date) = " & date2 & " and Doctor_ID = '" & cmb_doc_id & "'", cnPatients, adOpenDynamic, adLockReadOnly
Debug.Print "Record Count : " & rsrec.RecordCount

While rsrec.EOF = False
   i = i + 1
   Debug.Print "Appointment Month : " & Month(rsrec![Appointment_Date])
   rsrec.MoveNext
Wend
txt_num_of_apps = i
numapp = i

rsvis.Open "Select * from Visit_Details where Month(Visit_Date) = " & date2 & " and Doctor_ID = '" & cmb_doc_id & "'", cnPatients, adOpenDynamic, adLockReadOnly
Debug.Print "Record Count : " & rsvis.RecordCount

While rsvis.EOF = False
   j = j + 1
   Debug.Print "Visit Month : " & Month(rsvis![Visit_Date])
   rsvis.MoveNext
Wend
txt_num_of_visits = j
numvis = j
rsvis.Close
rsrec.Close
'Visit_ID      Visit_Time  Doctor_ID   Admission_ID    Patient_ID  Description

'rscat.Close
rsDocs.Open "select Doctor_CCharge,Doctor_VCharge from Doctor_Details where Doctor_ID= '" & cmb_doc_id.Text & "'", cnPatients, adOpenDynamic, adLockOptimistic
ch = rsDocs!Doctor_CCharge * numapp
vc = rsDocs!Doctor_VCharge * numvis
rsDocs.Close

End Sub



Private Sub cmd_back_Click()
frm_employee.Show
Unload Me
End Sub

Private Sub cmd_clear_Click()
Call clear
End Sub

Private Sub cmd_del_Click()

If cmb_doc_id = "" Then
MsgBox "There Is No Current Record", vbInformation
Else
res = MsgBox("Do You Want To Delete The Current Record ? ", vbCritical + vbYesNo, "Data Deletion")
If res = vbYes Then
cnPatients.Execute ("delete from Doctor_salary where decode='" & cmb_doc_id.Text & "'")
Call clear
ElseIf res = vbNo Then
MsgBox "Deletion Cancled", vbInformation
End If
End If

End Sub

Private Sub cmd_modify_Click()


If cmb_doc_id = "" Then
MsgBox "There Is No Current Record", vbInformation
Else
res = MsgBox("Do You Want To Modify The Current Record ? ", vbCritical + vbYesNo, "Data Modification")
If res = vbYes Then

cnPatients.Execute ("delete from Doctor_salary where decode='" & cmb_doc_id & "'")
cnPatients.Execute ("Insert into Doctor_salary values('" & cmb_doc_id & "','" & txt_doc_name & "','" & ch & "','" & vc & "','" & bp & "','" & da & "','" & hr & "','" & ca & "','" & ta & "','" & pf & "','" & ins & "','" & it & "','" & pt & "','" & gpa & "','" & deduct & "','" & Net & "')")

Call clear
ElseIf res = vbNo Then
MsgBox "Modifcation Cancled", vbInformation
End If
End If
End Sub

Private Sub cmd_save_Click()

On Error GoTo er1

cnPatients.Execute ("Insert into Doctor_salary values('" & cmb_doc_id & "','" & txt_doc_name & "','" & ch & "','" & vc & "','" & bp & "','" & da & "','" & hr & "','" & ca & "','" & ta & "','" & pf & "','" & ins & "','" & it & "','" & pt & "','" & gpa & "','" & deduct & "','" & Net & "')")

Exit Sub

er1:
Debug.Print er1

Call clear
'.cmb_doc_id = ""
'.cmb_month = ""
'.txt_doc_name = ""
'.txt_num_of_apps = ""
'.txt_num_of_visits = ""


End Sub

Private Sub cmd_view_doc_Click()
frm_permanantdoctorsalary.Show
End Sub

Private Sub Command1_Click()
frm_permanantdoctorsalary.Show
End Sub



Private Sub da_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
    End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 22 Then KeyAscii = 0: Exit Sub
    KeyAscii = DataEntryValidation(KeyAscii, ActiveControl.Tag)
End Sub


Private Sub Form_Load()
Dim rs As New ADODB.Recordset
'cnPatients.Open "Provider=Microsoft.Jet.OLEDB.4.0;data source=& \HMS.mdb;"
'cnPatients.CursorLocation = adUseClient
frm_employee.Show
Dim sa As String
sa = "Permanent Doctor"
rs.Open "select Distinct Doctor_ID, Doctor_Type,Doctor_FName,Doctor_LName from Doctor_Details where Doctor_Type= '" & sa & "'", cnPatients, adOpenDynamic, adLockOptimistic

rs.MoveFirst
While rs.EOF = False
cmb_doc_id.AddItem rs!Doctor_ID
rs.MoveNext
Wend

Frame4.Visible = False
cmb_month.Visible = False

'Doctor_ID
'Doctor_Details
'Doctor_Type
'Permanent Doctor
'Doctor_FName Doctor_LName


End Sub

Private Sub deduct_LostFocus()
    gpa.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) + Val(ch.Text) + Val(vc.Text) - Val(deduct.Text)
    deduct.Text = Val(pf.Text) + Val(ins.Text) + Val(it.Text) + Val(pt.Text)
    Net.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) - Val(pf.Text) - Val(ins.Text) - Val(it.Text) - Val(pt.Text)
End Sub
Private Sub gpa_LostFocus()
    gpa.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) + Val(ch.Text) + Val(vc.Text) - Val(deduct.Text)
    Net.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) - Val(pf.Text) - Val(ins.Text) - Val(it.Text) - Val(pt.Text)
End Sub

Private Sub hr_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
    End If
End Sub

Private Sub hr_LostFocus()
    gpa.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) + Val(ch.Text) + Val(vc.Text) - Val(deduct.Text)
    Net.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) - Val(pf.Text) - Val(ins.Text) - Val(it.Text) - Val(pt.Text)
End Sub
Private Sub ins_LostFocus()
    deduct.Text = Val(pf.Text) + Val(ins.Text) + Val(it.Text) + Val(pt.Text)
    Net.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) - Val(pf.Text) - Val(ins.Text) - Val(it.Text) - Val(pt.Text)
End Sub
Private Sub it_LostFocus()
    deduct.Text = Val(pf.Text) + Val(ins.Text) + Val(it.Text) + Val(pt.Text)
    Net.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) - Val(pf.Text) - Val(ins.Text) - Val(it.Text) - Val(pt.Text)
End Sub
Private Sub Net_LostFocus()
    gpa.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) + Val(ch.Text) + Val(vc.Text) - Val(deduct.Text)
    deduct.Text = Val(pf.Text) + Val(ins.Text) + Val(it.Text) + Val(pt.Text)
    Net.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) - Val(pf.Text) - Val(ins.Text) - Val(it.Text) - Val(pt.Text)
End Sub
Private Sub pf_LostFocus()
    deduct.Text = Val(pf.Text) + Val(ins.Text) + Val(it.Text) + Val(pt.Text)
    Net.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) - Val(pf.Text) - Val(ins.Text) - Val(it.Text) - Val(pt.Text)
End Sub
Private Sub pt_LostFocus()
    deduct.Text = Val(pf.Text) + Val(ins.Text) + Val(it.Text) + Val(pt.Text)
    Net.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) - Val(pf.Text) - Val(ins.Text) - Val(it.Text) - Val(pt.Text)
End Sub
Private Sub ta_LostFocus()
    gpa.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) + Val(ch.Text) + Val(vc.Text) - Val(deduct.Text)
    Net.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) - Val(pf.Text) - Val(ins.Text) - Val(it.Text) - Val(pt.Text)
End Sub



Public Sub clear()
With frm_app_count
.bp = ""
.ca = ""
.ch = ""
.da = ""
.deduct = ""
.gpa = ""
.hr = ""
.ins = ""
.it = ""
.Net = ""
.pf = ""
.pt = ""
.ta = ""
.vc = ""
End With
End Sub
