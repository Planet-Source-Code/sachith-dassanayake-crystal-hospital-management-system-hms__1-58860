VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmOPSerBillPayments 
   BackColor       =   &H00FF8080&
   Caption         =   "MEDICAL APPOINTMENT BILL PAYMENTS"
   ClientHeight    =   11355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11820
   Icon            =   "frmOPSerBillPayments.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   11355
   ScaleWidth      =   11820
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport SInvoice 
      Left            =   3240
      Top             =   10440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&SAVE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   3960
      Picture         =   "frmOPSerBillPayments.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Click To Save Bill Payment Information"
      Top             =   10200
      Width           =   1185
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&CLOSE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   6840
      Picture         =   "frmOPSerBillPayments.frx":5C90
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Click To Close"
      Top             =   10200
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print Invoice"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   5400
      Picture         =   "frmOPSerBillPayments.frx":6194
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   10200
      Width           =   1215
   End
   Begin VB.ComboBox cmbBillNo 
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
      Left            =   9600
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin VB.ComboBox cmbPatientID 
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
      Left            =   1800
      TabIndex        =   6
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtBillDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   9600
      Locked          =   -1  'True
      TabIndex        =   20
      ToolTipText     =   "Bill Date"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox txtBillNo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   9720
      Locked          =   -1  'True
      TabIndex        =   19
      ToolTipText     =   "Bill Date"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtPatientID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   18
      ToolTipText     =   "Bill Date"
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Payment Info"
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
      Height          =   2415
      Left            =   240
      TabIndex        =   7
      Top             =   7560
      Width           =   11505
      Begin VB.TextBox txtBillStatus 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   36
         ToolTipText     =   "Customer Bill Status"
         Top             =   1290
         Width           =   1755
      End
      Begin VB.TextBox txtBalAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   35
         ToolTipText     =   "Bill Balance Amount"
         Top             =   840
         Width           =   1755
      End
      Begin VB.TextBox txtPayingAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   2520
         TabIndex        =   34
         Tag             =   "Amt"
         ToolTipText     =   "Enter The Paying Amount"
         Top             =   360
         Width           =   1755
      End
      Begin VB.TextBox txtDDNo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   7680
         TabIndex        =   13
         ToolTipText     =   "Enter the DD Number"
         Top             =   840
         Width           =   3585
      End
      Begin VB.ComboBox cmbBank 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   7680
         TabIndex        =   12
         ToolTipText     =   "Select The Bank Name"
         Top             =   1650
         Width           =   3585
      End
      Begin VB.OptionButton optOthers 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "OTHERS"
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
         Left            =   10080
         TabIndex        =   11
         ToolTipText     =   "Click Here If Payment By Others"
         Top             =   420
         Width           =   1155
      End
      Begin VB.OptionButton optCheque 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "CHEQUE"
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
         Left            =   8880
         TabIndex        =   10
         ToolTipText     =   "Click Here If Payment By Cheque"
         Top             =   420
         Width           =   1095
      End
      Begin VB.OptionButton optDD 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "CREDIT CARD"
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
         Left            =   7080
         TabIndex        =   9
         ToolTipText     =   "Click Here If Payment By DD"
         Top             =   420
         Width           =   1665
      End
      Begin VB.OptionButton optCash 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "CASH"
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
         Left            =   6060
         TabIndex        =   8
         ToolTipText     =   "Click Here If Payment Is Cash"
         Top             =   420
         Value           =   -1  'True
         Width           =   825
      End
      Begin MSComCtl2.DTPicker dtpDDDate 
         Height          =   285
         Left            =   7680
         TabIndex        =   14
         Top             =   1200
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   503
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
         Format          =   45940737
         CurrentDate     =   38330
      End
      Begin MSComCtl2.DTPicker dtpPayDate 
         Height          =   315
         Left            =   2520
         TabIndex        =   37
         Top             =   1800
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
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
         Format          =   45940737
         CurrentDate     =   38330
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C96C59&
         BackStyle       =   0  'Transparent
         Caption         =   "Balance Amount :"
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
         Left            =   720
         TabIndex        =   41
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C96C59&
         BackStyle       =   0  'Transparent
         Caption         =   "Paying Amount :"
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
         Left            =   825
         TabIndex        =   40
         Top             =   360
         Width           =   1590
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C96C59&
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Date :"
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
         Left            =   930
         TabIndex        =   39
         Top             =   1800
         Width           =   1485
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C96C59&
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Status :"
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
         Left            =   1320
         TabIndex        =   38
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "CHECK NO :"
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
         Left            =   6450
         TabIndex        =   17
         Top             =   870
         Width           =   1065
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "DATE :"
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
         Left            =   6450
         TabIndex        =   16
         Top             =   1290
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "BANK :"
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
         Left            =   6450
         TabIndex        =   15
         Top             =   1710
         Width           =   630
      End
   End
   Begin VB.TextBox txtBillAmt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Bill Total Amount"
      Top             =   3180
      Width           =   2415
   End
   Begin VB.TextBox txtNetValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   9030
      Locked          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Bill Net Value"
      Top             =   3180
      Width           =   2415
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   5190
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Bill Discount"
      Top             =   3180
      Width           =   2415
   End
   Begin VB.TextBox txtBal 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   7020
      Locked          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Bill Balance Amount"
      Top             =   6600
      Width           =   2415
   End
   Begin VB.TextBox txtPaidAmt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Total Amount Paid"
      Top             =   6600
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid MFG 
      Height          =   1965
      Left            =   210
      TabIndex        =   21
      ToolTipText     =   "Bill Payments List"
      Top             =   4440
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   3466
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ForeColor       =   128
      ForeColorFixed  =   8388608
      GridColor       =   13200473
      GridColorFixed  =   13200473
      AllowUserResizing=   3
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      FormHeightDT    =   11865
      FormWidthDT     =   11940
      FormScaleHeightDT=   11355
      FormScaleWidthDT=   11820
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Discount :"
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
      Left            =   3960
      TabIndex        =   42
      Top             =   3240
      Width           =   960
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date :"
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
      Left            =   8520
      TabIndex        =   30
      Top             =   1095
      Width           =   930
   End
   Begin VB.Label lblManTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "OUT PATIENT MEDICAL APPOINTMENT BILL PAYMENTS"
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
      Left            =   405
      TabIndex        =   29
      Top             =   360
      Width           =   11175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Code :"
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
      Left            =   240
      TabIndex        =   28
      Top             =   1695
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Number :"
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
      Left            =   8280
      TabIndex        =   27
      Top             =   1710
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   30
      X2              =   11640
      Y1              =   2085
      Y2              =   2070
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   30
      X2              =   11640
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      Height          =   945
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   11565
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "BILL PAYMENT DETAILS :"
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
      Left            =   240
      TabIndex        =   26
      Top             =   4020
      Width           =   2385
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   11760
      Y1              =   7215
      Y2              =   7200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Amt :"
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
      Left            =   420
      TabIndex        =   25
      Top             =   3240
      Width           =   870
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Net Value :"
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
      Left            =   7920
      TabIndex        =   24
      Top             =   3240
      Width           =   1050
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Balance :"
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
      Left            =   6030
      TabIndex        =   23
      Top             =   6690
      Width           =   885
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount Paid :"
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
      Left            =   1440
      TabIndex        =   22
      Top             =   6690
      Width           =   1905
   End
End
Attribute VB_Name = "frmOPSerBillPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RowNo As Integer
Private Sub cmbBillNo_Click()
Dim ctl As Control

For Each ctl In Controls
    If TypeOf ctl Is TextBox And TypeOf ctl Is ComboBox Then
    ctl.Text = ""
    End If
Next
MFG.clear
Refresh_Data

Dim rsPay As Recordset
Set rsPay = New ADODB.Recordset
Dim rsBill As Recordset
Set rsBill = New ADODB.Recordset


Dim i As Integer
Dim s As Double


If cmbBillNo.Text = "" Then
    Exit Sub
End If

rsBill.Open "Select * from Service_Appointment_Bill where Appointment_Bill_ID = '" & cmbBillNo & "'", cnPatients, adOpenDynamic, adLockPessimistic

If rsBill.EOF = True Then
    rsBill.Close
Else
    txtBillDate.Text = Format(rsBill!Bill_Date, "dd-MMM-yyyy")
    'txtBillTerms.Text = rsBill!CreditYN
    txtBillAmt.Text = rsBill!Grand_Total
    txtDiscount.Text = rsBill!Discount
    txtNetValue.Text = rsBill!Net_Value
    
    'rsPay.open "Select count(*) from OPBillDetails where BillId=" & cmbBillNo.ItemData(cmbBillNo.ListIndex))
    'If rsPay.EOF = True Then
        'rsPay.Close
    'Else
        'txtBillItems.Text = rsPay(0)
        'rsPay.Close
    'End If
    rsBill.Close
End If

rsBill.Open "Select * from Service_Appointment_Bill_Payment where Appointment_Bill_ID='" & cmbBillNo & "'", cnPatients, adOpenDynamic, adLockPessimistic
If rsBill.EOF = True Then
    rsBill.Close
    txtPaidAmt.Text = "0"
    txtBal.Text = txtNetValue.Text
Else
    i = 1
    s = 0
    MFG.Rows = 2
    Do While rsBill.EOF = False
        MFG.TextMatrix(i, 0) = i
        MFG.TextMatrix(i, 1) = rsBill!Amount_Paid
        MFG.TextMatrix(i, 2) = Format(rsBill!Paid_Date, "dd-MMM-yyyy")
        MFG.TextMatrix(i, 3) = rsBill!Payment_Type
        If IsNull(rsBill!CheckNo) = False Then
            MFG.TextMatrix(i, 4) = rsBill!CheckNo
        End If
        If IsNull(rsBill!CheckDate) = False Then
            MFG.TextMatrix(i, 5) = rsBill!CheckDate
        End If
        If IsNull(rsBill!Bank) = False Then
            MFG.TextMatrix(i, 6) = rsBill!Bank
        End If
        s = s + Val(rsBill!Amount_Paid)
        rsBill.MoveNext
        i = i + 1
        MFG.Rows = MFG.Rows + 1
    Loop
    rsBill.Close
    txtPaidAmt.Text = s
    txtBal.Text = Round(Val(txtNetValue.Text) - Val(txtPaidAmt.Text), 2)
End If
RowNo = MFG.Rows - 1

txtPayingAmt.SetFocus













End Sub

Private Sub cmbPatientID_Click()
Dim rsBill As Recordset
Set rsBill = New ADODB.Recordset

Dim i As Integer
i = 0
If cmbPatientID.Text = "" Then
    Exit Sub
End If
cmbBillNo.clear

rsBill.Open "Select * from Service_Appointment_Bill where Patient_ID ='" & cmbPatientID.Text & "'", cnPatients, adOpenDynamic, adLockPessimistic
If rsBill.EOF = True Then
    rsBill.Close
Else
    Do While rsBill.EOF = False
        cmbBillNo.AddItem (rsBill(0))
        rsBill.MoveNext
        i = i + 1
    Loop
    rsBill.Close
End If

End Sub

Private Sub cmdsave_click()
Dim str1 As String
Dim BillPayID As String
Dim rsAddBill As New Recordset


If txtPayingAmt.Text = "" Then
    MsgBox "Paying Amount Not Found...", vbCritical + vbOKOnly
    txtPayingAmt.SetFocus
    Exit Sub
End If

If optCash.Value = True Then
    str1 = "CASH"
ElseIf optDD.Value = True Then
    str1 = "Credit Card"
ElseIf optCheque.Value = True Then
    str1 = "Cheque"
Else
    str1 = "Others"
End If


    Set rsAddBill = New ADODB.Recordset
    BillPayID = Functions.UID(6, "PayID_")

    rsAddBill.Open "Select * from Service_Appointment_Bill_Payment", cnPatients, adOpenKeyset, adLockPessimistic
      
    While rsAddBill.EOF = False
        If rsAddBill(0) = BillPayID Then
            BillPayID = Functions.UID(6, "PayID_")
            rsAddBill.MoveFirst
            
        Else
            
        End If
      
    rsAddBill.MoveNext
    
    Wend




If MsgBox("Confirm To Save Bill Information ?", vbQuestion + vbYesNo) = vbYes Then
   
    'cnPatients.BeginTrans
    
    If optCash.Value = True Then
        cnPatients.Execute ("Insert into Service_Appointment_Bill_Payment values('" & BillPayID & "','" & cmbBillNo & "'," & Val(txtPayingAmt.Text) & ",'" & Format(dtpPayDate.Value, "mm/dd/yy") & "','" & str1 & "',Null,Null,Null)")
        MFG.TextMatrix(RowNo, 0) = RowNo
        MFG.TextMatrix(RowNo, 1) = Val(txtPayingAmt.Text)
        MFG.TextMatrix(RowNo, 2) = Format(dtpPayDate.Value, "dd-MMM-yyyy")
        MFG.TextMatrix(RowNo, 3) = str1
        RowNo = RowNo + 1
        MFG.Rows = MFG.Rows + 1
    Else
        If txtDDNo.Text = "" Or cmbBank.Text = "" Then
            MsgBox "Check Number or Bank Name Not Found...", vbCritical + vbOKOnly
            txtDDNo.SetFocus
            Exit Sub
        End If
        Debug.Print BillPayID
        Debug.Print cmbBillNo
        Debug.Print txtPayingAmt
        Debug.Print dtpPayDate
        Debug.Print str1
        Debug.Print txtDDNo
        Debug.Print dtpDDDate
        Debug.Print cmbBank
        
        
        cnPatients.Execute ("Insert into Service_Appointment_Bill_Payment values('" & BillPayID & "','" & cmbBillNo & "'," & Val(txtPayingAmt.Text) & ",#" & Format(dtpPayDate.Value, "short date") & "#,'" & str1 & "','" & txtDDNo.Text & "',#" & Format(dtpDDDate.Value, "short date") & "#,'" & cmbBank.Text & "')")
        MFG.TextMatrix(RowNo, 0) = RowNo
        MFG.TextMatrix(RowNo, 1) = Val(txtPayingAmt.Text)
        MFG.TextMatrix(RowNo, 2) = Format(dtpPayDate.Value, "dd-MMM-yyyy")
        MFG.TextMatrix(RowNo, 3) = str1
        MFG.TextMatrix(RowNo, 4) = txtDDNo.Text
        MFG.TextMatrix(RowNo, 5) = Format(dtpDDDate.Value, "dd-MMM-yyyy")
        MFG.TextMatrix(RowNo, 6) = cmbBank.Text
        RowNo = RowNo + 1
        MFG.Rows = MFG.Rows + 1
    End If
   Call Txt_Clear
    txtPayingAmt.SetFocus
    'cnPatients.CommitTrans


End If















End Sub

Private Sub Command1_Click()
Dim strReport As String
Dim strTXT As String
strTXT = cmbBillNo.Text
strReport = App.Path & "\Reports\OPSerInvoice.rpt"


SInvoice.ReportFileName = App.Path & "\Reports\OPSerInvoice.rpt"
SInvoice.DiscardSavedData = True
SInvoice.ReplaceSelectionFormula ("{Service_Appointment_Bill_Payment.Appointment_Bill_ID} = '" & strTXT & "'")


SInvoice.WindowState = crptMaximized
SInvoice.Action = 1




End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 22 Then KeyAscii = 0: Exit Sub
    KeyAscii = DataEntryValidation(KeyAscii, ActiveControl.Tag)
End Sub
Private Sub Form_Load()
    Call Functions.DisableMenu
    Me.WindowState = vbMaximized
    Dim i As Integer
Dim rsBill As Recordset

Set rsBill = New ADODB.Recordset
i = 0
rsBill.Open "Select * from Patient_Details where Patient_ID in (Select Distinct Patient_ID from Service_Appointment_Bill)", cnPatients, adOpenDynamic, adLockPessimistic
If rsBill.EOF = True Then
    rsBill.Close
Else
    Do While rsBill.EOF = False
        cmbPatientID.AddItem (rsBill(0))
        rsBill.MoveNext
        i = i + 1
    Loop
    rsBill.Close
End If
    
    
    
    
    
    Refresh_Data
    RowNo = 0
End Sub


Private Sub Refresh_Data()

dtpPayDate = Date

MFG.clear
MFG.ColWidth(0) = 1000
MFG.ColAlignment(0) = 4
For i = 1 To 6 Step 1
    MFG.ColWidth(i) = 2000
    MFG.ColAlignment(i) = 4
Next i
MFG.TextMatrix(0, 0) = "NO"
MFG.TextMatrix(0, 1) = "AMOUNT PAID"
MFG.TextMatrix(0, 2) = "PAID DATE"
MFG.TextMatrix(0, 3) = "PAY TYPE"
MFG.TextMatrix(0, 4) = "CREDIT CARD/CHEQUE NO"
MFG.TextMatrix(0, 5) = "CHEQUE DATE"
MFG.TextMatrix(0, 6) = "BANK"
End Sub


Private Sub Txt_Clear()
Dim i As Integer
Dim s As Double
s = 0
txtPayingAmt.Text = ""
txtBalAmt.Text = ""
'txtBalAdv.Text = ""
txtBillStatus.Text = ""
txtDDNo.Text = ""

For i = 1 To MFG.Rows - 2 Step 1
  s = s + Val(MFG.TextMatrix(i, 1))
  Debug.Print MFG.TextMatrix(i, 1)
Next i
txtPaidAmt.Text = s
txtBalAmt.Text = Round(Val(txtBal.Text) - Val(txtPaidAmt.Text), 2)
txtBal.Text = Round(Val(txtNetValue.Text) - Val(txtPaidAmt.Text), 2)
txtBalAmt = ""
End Sub



Private Sub Form_Unload(Cancel As Integer)
Call Functions.EnableMenu
End Sub


Private Sub txtPayingAmt_LostFocus()

If txtPayingAmt.Text <> "" Then
    If Val(txtPayingAmt.Text) = 0 Then
        MsgBox "Paying Amount Cannot Be Zero...", vbInformation + vbOKOnly
        txtPayingAmt.SetFocus
        Exit Sub
    End If
    If Val(txtPayingAmt.Text) > Val(txtBal.Text) Then
        MsgBox "Paying Amount Cannot Be Greater Than Balance Amount...", vbCritical + vbOKOnly
        txtPayingAmt.Text = ""
        txtPayingAmt.SetFocus
        Exit Sub
    End If
     txtBalAmt.Text = Round((Val(txtBal.Text) - Val(txtPayingAmt.Text)), 2)
     'If Val(txtCustomerAdv.Text) <> 0 Then
        'If Val(txtCustomerAdv.Text) > Val(txtBalAmt.Text) Then
            'txtCustomerAdv.Text = Val(txtCustomerAdv.Text) - Val(txtBalAmt.Text)
            'txtBalAdv.Text = Val(txtCustomerAdv.Text) - Val(txtBalAmt.Text)
            'txtBalAmt.Text = "0"
        'Else
            'txtBalAmt.Text = Round(Val(txtBalAmt.Text) - Val(txtCustomerAdv.Text), 2)
            'txtCustomerAdv.Text = "0"
            'txtBalAdv.Text = "0"
        'End If
    'Else
        'txtBalAdv.Text = "0"
    'End If
    If Val(txtBalAmt.Text) = 0 Then
        txtBillStatus.Text = "Paid"
    Else
        txtBillStatus.Text = "Un-Paid"
    End If


End If

End Sub

