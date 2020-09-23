VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmIPBillPayments 
   BackColor       =   &H00FF8080&
   Caption         =   "In Patient Bill Payments"
   ClientHeight    =   11520
   ClientLeft      =   1860
   ClientTop       =   645
   ClientWidth     =   12225
   Icon            =   "frmIPBillPayments.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   12225
   WindowState     =   2  'Maximized
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
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   42
      ToolTipText     =   "Bill Date"
      Top             =   1560
      Width           =   2055
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
      Left            =   5280
      Picture         =   "frmIPBillPayments.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   10320
      Width           =   1215
   End
   Begin VB.ComboBox cmbAdmitID 
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
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text2 
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
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   30
      ToolTipText     =   "Bill Date"
      Top             =   1560
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
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   19
      ToolTipText     =   "Bill Date"
      Top             =   960
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
      Height          =   2295
      Left            =   600
      TabIndex        =   7
      Top             =   7560
      Width           =   11265
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         Left            =   6870
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
         Left            =   6870
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
         Left            =   9720
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
         Left            =   8520
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
         Left            =   6720
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
         Left            =   5700
         TabIndex        =   8
         ToolTipText     =   "Click Here If Payment Is Cash"
         Top             =   420
         Value           =   -1  'True
         Width           =   825
      End
      Begin MSComCtl2.DTPicker dtpDDDate 
         Height          =   285
         Left            =   6870
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
         Format          =   47841281
         CurrentDate     =   38330
      End
      Begin MSComCtl2.DTPicker dtpPayDate 
         Height          =   315
         Left            =   2520
         TabIndex        =   35
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
         Format          =   47841281
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
         Top             =   1320
         Width           =   1095
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
         Left            =   6090
         TabIndex        =   17
         Top             =   870
         Width           =   705
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
         Left            =   6090
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
         Left            =   6090
         TabIndex        =   15
         Top             =   1710
         Width           =   630
      End
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
      Left            =   3840
      Picture         =   "frmIPBillPayments.frx":5C62
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Click To Save Bill Payment Information"
      Top             =   10320
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
      Left            =   6720
      Picture         =   "frmIPBillPayments.frx":6110
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Click To Close"
      Top             =   10320
      Width           =   1185
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
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Bill Total Amount"
      Top             =   3060
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
      Left            =   9270
      Locked          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Bill Net Value"
      Top             =   3060
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
      Left            =   5430
      Locked          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Bill Discount"
      Top             =   3060
      Width           =   1695
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
      Left            =   7860
      Locked          =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "Bill Balance Amount"
      Top             =   6480
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
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Total Amount Paid"
      Top             =   6480
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid MFG 
      Height          =   1965
      Left            =   450
      TabIndex        =   20
      ToolTipText     =   "Bill Payments List"
      Top             =   4200
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
   Begin VB.TextBox txtPayNo 
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
      Left            =   9960
      Locked          =   -1  'True
      TabIndex        =   18
      ToolTipText     =   "Bill Date"
      Top             =   1560
      Width           =   1935
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
      FormHeightDT    =   12030
      FormWidthDT     =   12345
      FormScaleHeightDT=   11520
      FormScaleWidthDT=   12225
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   7200
      TabIndex        =   44
      Top             =   3120
      Width           =   210
   End
   Begin VB.Label Label3 
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
      Left            =   4800
      TabIndex        =   43
      Top             =   1560
      Width           =   1245
   End
   Begin VB.Label Label1 
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
      Left            =   4320
      TabIndex        =   41
      Top             =   3120
      Width           =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Admission ID :"
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
      Left            =   480
      TabIndex        =   31
      Top             =   1560
      Width           =   1410
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
      Left            =   8880
      TabIndex        =   28
      Top             =   975
      Width           =   930
   End
   Begin VB.Label lblManTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "IN PATIENT BILL PAYMENTS"
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
      Left            =   3480
      TabIndex        =   27
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice No:"
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
      TabIndex        =   26
      Top             =   1590
      Width           =   1110
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   270
      X2              =   11880
      Y1              =   1965
      Y2              =   1950
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   270
      X2              =   11880
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      Height          =   1305
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
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
      Left            =   480
      TabIndex        =   25
      Top             =   3900
      Width           =   2385
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   360
      X2              =   12000
      Y1              =   7095
      Y2              =   7080
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
      Left            =   660
      TabIndex        =   24
      Top             =   3120
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
      Left            =   8160
      TabIndex        =   23
      Top             =   3120
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
      Left            =   6870
      TabIndex        =   22
      Top             =   6570
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
      Left            =   2280
      TabIndex        =   21
      Top             =   6570
      Width           =   1905
   End
End
Attribute VB_Name = "frmIPBillPayments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RowNo As Integer




Private Sub cmbAdmitID_Click()
Dim rsAddBillID As Recordset
Set rsAddBillID = New ADODB.Recordset
Dim rsBill As Recordset
Set rsBill = New ADODB.Recordset

Dim i As Integer
Dim s As Double

rsAddBillID.Open "select * from Patient_Bill Where Admission_ID = '" & cmbAdmitID & "'", cnPatients, adOpenDynamic, adLockReadOnly

If rsAddBillID.EOF = True Then
    MsgBox "Cannot Pay the bill unitl patient discharge", vbCritical, "An Error Occured"
    rsAddBillID.Close
    Exit Sub
ElseIf rsAddBillID.RecordCount = 1 Then
    txtBillNo = rsAddBillID(0)
    txtNetValue = rsAddBillID![Net_Value]
    txtDiscount = rsAddBillID![Discount]
    
    txtBillAmt = Val(txtNetValue) + (Val(txtNetValue) * Val(txtDiscount))
    
    
rsBill.Open "Select * from Patient_Bill_Payment where PatientBill_ID='" & txtBillNo & "'", cnPatients, adOpenDynamic, adLockPessimistic
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

    
    
    
    
    
    
    
    
    
    
    
    
    
   
ElseIf rsAddBillID.RecordCount > 1 Then
    MsgBox "Error", vbCritical
    rsAddBillID.Close
    Exit Sub
End If



End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
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
    BillPayID = Functions.UID(6, "IPBPay_")

    rsAddBill.Open "Select * from Patient_Bill_Payment", cnPatients, adOpenKeyset, adLockPessimistic
      
    While rsAddBill.EOF = False
        If rsAddBill(0) = BillPayID Then
            BillPayID = Functions.UID(6, "IPBPay_")
            rsAddBill.MoveFirst
            
        Else
            
        End If
      
    rsAddBill.MoveNext
    
    Wend




If MsgBox("Confirm To Save Bill Information ?", vbQuestion + vbYesNo) = vbYes Then
   
    'cnPatients.BeginTrans
    
    If optCash.Value = True Then
        cnPatients.Execute ("Insert into Patient_Bill_Payment values('" & BillPayID & "','" & txtBillNo & "'," & Val(txtPayingAmt.Text) & ",'" & Format(dtpPayDate.Value, "mm/dd/yy") & "','" & str1 & "',Null,Null,Null)")
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
        
        
        cnPatients.Execute ("Insert into Patient_Bill_Payment values('" & BillPayID & "','" & txtBillNo & "'," & Val(txtPayingAmt.Text) & ",#" & Format(dtpPayDate.Value, "short date") & "#,'" & str1 & "','" & txtDDNo.Text & "',#" & Format(dtpDDDate.Value, "short date") & "#,'" & cmbBank.Text & "')")
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 22 Then KeyAscii = 0: Exit Sub
    KeyAscii = DataEntryValidation(KeyAscii, ActiveControl.Tag)
End Sub

Private Sub Form_Load()
Call Functions.DisableMenu
dtpPayDate = Date



txtBillDate = Date
Call addPatientID
Call GenerateBillID
Call FinalBillAmount
Refresh_Data
RowNo = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Functions.EnableMenu
End Sub

Private Sub GenerateBillID()
    Dim rsAddPatient As Recordset
    Dim PID As String
    Set rsAddPatient = New ADODB.Recordset
  
    PID = Functions.UID(6, "IPBillID_")
    rsAddPatient.Open " Select * from Patient_Bill_Payment", cnPatients, adOpenDynamic, adLockPessimistic
    While rsAddPatient.EOF = False
        If rsAddPatient(0) = PID Then
            ID = True
            PID = Functions.UID(6, "IPBillID_")
            rsAddPatient.MoveFirst
        Else
            ID = False
        End If
    rsAddPatient.MoveNext
    Wend
    rsAddPatient.Close
    txtPayNo = PID


End Sub

Private Sub FinalBillAmount()
Debug.Print "as"
End Sub

Private Sub addPatientID()

Dim rsAddPatient As Recordset
Set rsAddPatient = New ADODB.Recordset

rsAddPatient.Open "Select * from Admission_Details", cnPatients, adOpenDynamic, adLockReadOnly

While rsAddPatient.EOF = False
cmbAdmitID.AddItem (rsAddPatient(0))
rsAddPatient.MoveNext

Wend

rsAddPatient.Close
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
 
    If Val(txtBalAmt.Text) = 0 Then
        txtBillStatus.Text = "Paid"
    Else
        txtBillStatus.Text = "Un-Paid"
    End If
End If
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
