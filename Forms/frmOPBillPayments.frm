VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmOPBillPayments 
   BackColor       =   &H00FF8080&
   Caption         =   "Out Patient Bill Payments"
   ClientHeight    =   11505
   ClientLeft      =   1965
   ClientTop       =   765
   ClientWidth     =   11940
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOPBillPayments.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11505
   ScaleWidth      =   11940
   WindowState     =   2  'Maximized
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
      Picture         =   "frmOPBillPayments.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   10200
      Width           =   1215
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
      Picture         =   "frmOPBillPayments.frx":5C62
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Click To Close"
      Top             =   10200
      Width           =   1185
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
      Picture         =   "frmOPBillPayments.frx":6166
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Click To Save Bill Payment Information"
      Top             =   10200
      Width           =   1185
   End
   Begin VB.CommandButton cmdViewPatient 
      Caption         =   "..."
      Height          =   255
      Left            =   3840
      TabIndex        =   31
      Top             =   1920
      Width           =   495
   End
   Begin Crystal.CrystalReport DInvoice 
      Left            =   4920
      Top             =   9240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowProgressCtls=   0   'False
      WindowShowSearchBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
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
      Left            =   4200
      Locked          =   -1  'True
      TabIndex        =   28
      ToolTipText     =   "Total Amount Paid"
      Top             =   6720
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
      Left            =   7740
      Locked          =   -1  'True
      TabIndex        =   27
      ToolTipText     =   "Bill Balance Amount"
      Top             =   6720
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
      TabIndex        =   24
      ToolTipText     =   "Bill Discount"
      Top             =   3420
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
      TabIndex        =   23
      ToolTipText     =   "Bill Net Value"
      Top             =   3420
      Width           =   2415
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
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   22
      ToolTipText     =   "Bill Total Amount"
      Top             =   3420
      Width           =   2415
   End
   Begin VB.ComboBox cmbBillNo 
      Height          =   315
      Left            =   9600
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   1920
      Width           =   2055
   End
   Begin VB.ComboBox cmbPatientID 
      Height          =   315
      Left            =   1680
      TabIndex        =   20
      Top             =   1920
      Width           =   2055
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
      Left            =   600
      TabIndex        =   3
      Top             =   7560
      Width           =   11025
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
         Left            =   2040
         TabIndex        =   37
         Tag             =   "Amt"
         ToolTipText     =   "Enter The Paying Amount"
         Top             =   360
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   36
         ToolTipText     =   "Bill Balance Amount"
         Top             =   840
         Width           =   1755
      End
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
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   35
         ToolTipText     =   "Customer Bill Status"
         Top             =   1290
         Width           =   1755
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
         Left            =   5460
         TabIndex        =   9
         ToolTipText     =   "Click Here If Payment Is Cash"
         Top             =   420
         Value           =   -1  'True
         Width           =   825
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
         Left            =   6480
         TabIndex        =   8
         ToolTipText     =   "Click Here If Payment By DD"
         Top             =   420
         Width           =   1665
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
         Left            =   8280
         TabIndex        =   7
         ToolTipText     =   "Click Here If Payment By Cheque"
         Top             =   420
         Width           =   1095
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
         Left            =   9480
         TabIndex        =   6
         ToolTipText     =   "Click Here If Payment By Others"
         Top             =   420
         Width           =   1155
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
         Left            =   6630
         TabIndex        =   5
         ToolTipText     =   "Select The Bank Name"
         Top             =   1650
         Width           =   3585
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
         Left            =   6630
         TabIndex        =   4
         ToolTipText     =   "Enter the DD Number"
         Top             =   840
         Width           =   3585
      End
      Begin MSComCtl2.DTPicker dtpDDDate 
         Height          =   285
         Left            =   6630
         TabIndex        =   10
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
         Left            =   2040
         TabIndex        =   38
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
         Left            =   840
         TabIndex        =   42
         Top             =   1320
         Width           =   1095
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
         Left            =   450
         TabIndex        =   41
         Top             =   1800
         Width           =   1485
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
         Left            =   345
         TabIndex        =   40
         Top             =   360
         Width           =   1590
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
         Left            =   240
         TabIndex        =   39
         Top             =   840
         Width           =   1695
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
         Left            =   5850
         TabIndex        =   13
         Top             =   1710
         Width           =   630
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
         Left            =   5850
         TabIndex        =   12
         Top             =   1290
         Width           =   630
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "DD No :"
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
         Left            =   5850
         TabIndex        =   11
         Top             =   870
         Width           =   705
      End
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Bill Date"
      Top             =   1920
      Width           =   1935
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
      TabIndex        =   1
      ToolTipText     =   "Bill Date"
      Top             =   1920
      Width           =   1935
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
      TabIndex        =   0
      ToolTipText     =   "Bill Date"
      Top             =   1320
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid MFG 
      Height          =   1965
      Left            =   210
      TabIndex        =   14
      ToolTipText     =   "Bill Payments List"
      Top             =   4560
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
      Left            =   240
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      FormHeightDT    =   12015
      FormWidthDT     =   12060
      FormScaleHeightDT=   11505
      FormScaleWidthDT=   11940
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
      TabIndex        =   43
      Top             =   3480
      Width           =   960
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
      Left            =   2160
      TabIndex        =   30
      Top             =   6810
      Width           =   1905
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
      Left            =   6750
      TabIndex        =   29
      Top             =   6810
      Width           =   885
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
      TabIndex        =   26
      Top             =   3480
      Width           =   1050
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
      Top             =   3480
      Width           =   870
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   11760
      Y1              =   7335
      Y2              =   7320
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
      TabIndex        =   19
      Top             =   4260
      Width           =   2385
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      Height          =   945
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3120
      Width           =   11565
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   30
      X2              =   11640
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   30
      X2              =   11640
      Y1              =   2325
      Y2              =   2310
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
      TabIndex        =   18
      Top             =   1950
      Width           =   1245
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
      TabIndex        =   17
      Top             =   1935
      Width           =   1350
   End
   Begin VB.Label lblManTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "OUT PATIENT DOCTOR APPOINTMENT BILL PAYMENTS"
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
      Left            =   495
      TabIndex        =   16
      Top             =   240
      Width           =   10995
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
      TabIndex        =   15
      Top             =   1335
      Width           =   930
   End
End
Attribute VB_Name = "frmOPBillPayments"
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

rsBill.Open "Select * from Appointment_Bill where Appointment_Bill_ID = '" & cmbBillNo & "'", cnPatients, adOpenDynamic, adLockPessimistic

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

rsBill.Open "Select * from Appointment_Bill_Payment where Appointment_Bill_ID='" & cmbBillNo & "'", cnPatients, adOpenDynamic, adLockPessimistic
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
            MFG.TextMatrix(i, 4) = rsBill!DDNo
        End If
        If IsNull(rsBill!CheckDate) = False Then
            MFG.TextMatrix(i, 5) = rsBill!DDDate
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

rsBill.Open "Select * from Appointment_Bill where Patient_ID ='" & cmbPatientID.Text & "'", cnPatients, adOpenDynamic, adLockPessimistic
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

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdsave_click()
Dim str As String
Dim BillPayID As String
Dim rsAddBill As New Recordset


If txtPayingAmt.Text = "" Then
    MsgBox "Paying Amount Not Found...", vbCritical + vbOKOnly
    txtPayingAmt.SetFocus
    Exit Sub
End If

If optCash.Value = True Then
    str = "CASH"
ElseIf optDD.Value = True Then
    str = "DD"
ElseIf optCheque.Value = True Then
    str = "Cheque"
Else
    str = "Others"
End If


    Set rsAddBill = New ADODB.Recordset
    BillPayID = Functions.UID(6, "PayID_")

    rsAddBill.Open "Select * from Appointment_Bill_Payment", cnPatients, adOpenKeyset, adLockPessimistic
      
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
        cnPatients.Execute ("Insert into Appointment_Bill_Payment values('" & BillPayID & "','" & cmbBillNo & "'," & Val(txtPayingAmt.Text) & ",'" & Format(dtpPayDate.Value, "mm/dd/yy") & "','" & str & "',Null,Null,Null)")
        MFG.TextMatrix(RowNo, 0) = RowNo
        MFG.TextMatrix(RowNo, 1) = Val(txtPayingAmt.Text)
        MFG.TextMatrix(RowNo, 2) = Format(dtpPayDate.Value, "dd-MMM-yyyy")
        MFG.TextMatrix(RowNo, 3) = str
        RowNo = RowNo + 1
        MFG.Rows = MFG.Rows + 1
    Else
        If txtDDNo.Text = "" Or cmbBank.Text = "" Then
            MsgBox "DD Number or Bank Name Not Found...", vbCritical + vbOKOnly
            txtDDNo.SetFocus
            Exit Sub
        End If
        con.Execute ("Insert into Appointment_Bill_Payment values(" & BillPayID & "," & cmbBillNo & "," & Val(txtPayingAmt.Text) & ",'" & Format(dtpPayDate.Value, "mm/dd/yy") & "','" & str & "','" & txtDDNo.Text & "','" & Format(dtpDDDate.Value, "mm/dd/yy") & "','" & cmbBank.Text & "')")
        MFG.TextMatrix(RowNo, 0) = RowNo
        MFG.TextMatrix(RowNo, 1) = Val(txtPayingAmt.Text)
        MFG.TextMatrix(RowNo, 2) = Format(dtpPayDate.Value, "dd-MMM-yyyy")
        MFG.TextMatrix(RowNo, 3) = str
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

Private Sub cmdViewPatient_Click()
frmDisplayOutPatient.Show
End Sub

Private Sub Command1_Click()

Dim strReport As String
Dim strTXT As String
strTXT = cmbBillNo.Text
strReport = App.Path & "\Reports\OPDInvoice.rpt"


DInvoice.ReportFileName = App.Path & "\Reports\OPDInvoice.rpt"
DInvoice.DiscardSavedData = True
DInvoice.ReplaceSelectionFormula ("{Appointment_Bill_Payment.Appointment_Bill_ID} = '" & strTXT & "'")


DInvoice.WindowState = crptMaximized
DInvoice.Action = 1

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 22 Then KeyAscii = 0: Exit Sub
    KeyAscii = DataEntryValidation(KeyAscii, ActiveControl.Tag)
End Sub
Private Sub Form_Load()
    Call Functions.DisableMenu
    Me.WindowState = vbMaximized
    
    Dim rsBill As Recordset

Set rsBill = New ADODB.Recordset
i = 0
cmbPatientID.clear

rsBill.Open "Select * from Patient_Details where Patient_ID in (Select Distinct Patient_ID from Appointment_Bill)", cnPatients, adOpenDynamic, adLockPessimistic
If rsBill.EOF = True Then
    rsBill.Close
Else
    Do While rsBill.EOF = False
        cmbPatientID.AddItem (rsBill(0))
        rsBill.MoveNext
    Loop
    rsBill.Close
End If
    
    
    
    
    
    
    Refresh_Data
    RowNo = 0
End Sub


Private Sub Refresh_Data()

dtpPayDate = Date
Dim i As Integer

i = 0
MFG.clear
MFG.ColWidth(0) = 1000
MFG.ColAlignment(0) = 4
For i = 1 To 6 Step 1
    MFG.ColWidth(i) = 2000
    MFG.ColAlignment(i) = 4
Next i
MFG.TextMatrix(0, 0) = "SL NO"
MFG.TextMatrix(0, 1) = "AMOUNT PAID"
MFG.TextMatrix(0, 2) = "PAID DATE"
MFG.TextMatrix(0, 3) = "PAY TYPE"
MFG.TextMatrix(0, 4) = "DD/CHEQUE NO"
MFG.TextMatrix(0, 5) = "DD DATE"
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
