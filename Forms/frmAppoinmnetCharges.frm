VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmAppoinmnetCharges 
   BackColor       =   &H00FF8080&
   Caption         =   "Appoinment Invoice "
   ClientHeight    =   10395
   ClientLeft      =   870
   ClientTop       =   1260
   ClientWidth     =   12585
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAppoinmnetCharges.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10395
   ScaleWidth      =   12585
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   2175
      Left            =   480
      TabIndex        =   33
      Top             =   6600
      Width           =   5775
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
         Top             =   1170
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
         Top             =   720
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
         Tag             =   "amt"
         ToolTipText     =   "Enter The Paying Amount"
         Top             =   240
         Width           =   1755
      End
      Begin MSComCtl2.DTPicker dtpPayDate 
         Height          =   315
         Left            =   2520
         TabIndex        =   37
         Top             =   1680
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
         Format          =   20578305
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
         Top             =   720
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
         Top             =   330
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
         Top             =   1680
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
         Top             =   1200
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport DocInvoice 
      Left            =   1200
      Top             =   9360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&PRINT INVOICE"
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
      Left            =   5880
      Picture         =   "frmAppoinmnetCharges.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   9120
      Width           =   1335
   End
   Begin VB.TextBox txtNetValue 
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
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   29
      ToolTipText     =   "Bill Net Value"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txtHospitalCharges 
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
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   28
      ToolTipText     =   "Bill Total Amount"
      Top             =   2640
      Width           =   1575
   End
   Begin VB.TextBox txtDorS_Amount 
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
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   26
      ToolTipText     =   "Bill Total Amount"
      Top             =   2640
      Width           =   1575
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
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   24
      ToolTipText     =   "Bill Date"
      Top             =   960
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
      Left            =   10200
      Locked          =   -1  'True
      TabIndex        =   23
      ToolTipText     =   "Bill Date"
      Top             =   1560
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
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   22
      ToolTipText     =   "Bill Date"
      Top             =   1560
      Width           =   1935
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
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   14
      ToolTipText     =   "Bill Discount"
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox txtBillAmt 
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
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   13
      ToolTipText     =   "Bill Total Amount"
      Top             =   3120
      Width           =   1575
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
      Height          =   2175
      Left            =   6600
      TabIndex        =   2
      Top             =   6600
      Width           =   5625
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
         Left            =   1470
         TabIndex        =   8
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
         Left            =   1470
         TabIndex        =   7
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
         Left            =   4320
         TabIndex        =   6
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
         Left            =   3120
         TabIndex        =   5
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
         Left            =   1320
         TabIndex        =   4
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
         Left            =   300
         TabIndex        =   3
         ToolTipText     =   "Click Here If Payment Is Cash"
         Top             =   420
         Value           =   -1  'True
         Width           =   825
      End
      Begin MSComCtl2.DTPicker dtpDDDate 
         Height          =   285
         Left            =   1470
         TabIndex        =   21
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
         Format          =   20578305
         CurrentDate     =   38330
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
         Left            =   690
         TabIndex        =   11
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
         Left            =   690
         TabIndex        =   10
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
         Left            =   690
         TabIndex        =   9
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
      Left            =   4320
      Picture         =   "frmAppoinmnetCharges.frx":5C62
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Click To Save Bill Payment Information"
      Top             =   9120
      Width           =   1305
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
      Left            =   7440
      Picture         =   "frmAppoinmnetCharges.frx":6110
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Click To Close"
      Top             =   9120
      Width           =   1305
   End
   Begin MSFlexGridLib.MSFlexGrid MFG 
      Height          =   1965
      Left            =   690
      TabIndex        =   12
      ToolTipText     =   "Bill Payments List"
      Top             =   4170
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
      FormHeightDT    =   10905
      FormWidthDT     =   12705
      FormScaleHeightDT=   10395
      FormScaleWidthDT=   12585
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
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
      Left            =   9360
      TabIndex        =   31
      Top             =   3180
      Width           =   1050
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Hospital Charges : "
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
      Left            =   8580
      TabIndex        =   30
      Top             =   2640
      Width           =   1830
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor / Service Charges :"
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
      Left            =   915
      TabIndex        =   27
      Top             =   2640
      Width           =   2595
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
      Left            =   9195
      TabIndex        =   25
      Top             =   975
      Width           =   930
   End
   Begin VB.Label lblManTitle 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "OUT PATIENT BILL PAYMENTS"
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
      Left            =   3420
      TabIndex        =   20
      Top             =   240
      Width           =   6105
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
      Left            =   720
      TabIndex        =   19
      Top             =   1560
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
      Left            =   8880
      TabIndex        =   18
      Top             =   1590
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   510
      X2              =   12120
      Y1              =   1965
      Y2              =   1950
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   510
      X2              =   12120
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Label Label4 
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
      Left            =   6120
      TabIndex        =   17
      Top             =   3180
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C96C59&
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Amount :"
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
      TabIndex        =   16
      Top             =   3180
      Width           =   1230
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      Height          =   1305
      Left            =   600
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
      Left            =   720
      TabIndex        =   15
      Top             =   3900
      Width           =   2385
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0FFFF&
      BorderWidth     =   2
      X1              =   480
      X2              =   12120
      Y1              =   6375
      Y2              =   6360
   End
End
Attribute VB_Name = "frmAppoinmnetCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public Function setFormData()
    
    dtpPayDate.Value = Now
    dtpDDDate.Value = Now
    txtBillDate = Format(Date, "MM/DD/YYYY")
    txtPatientID.Text = BillPatientID
    txtBillNo.Text = AppBillID
    txtDorS_Amount.Text = strAmount
    txtHospitalCharges.Text = HospitalCharge
    txtDiscount.Text = Discount
    
    txtBillAmt = GrandTotal
    txtNetValue = Val(txtBillAmt) - Val(txtDiscount)
    
    
    txtDDNo.Text = ""
    txtDDNo.Locked = True
    cmbBank.Locked = True
    dtpDDDate.Enabled = False
    
End Function

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()

Dim strReport As String
Dim strTXT As String
strTXT = txtBillNo.Text
strReport = App.Path & "\Reports\OPDocInvoice.rpt"

DocInvoice.ReportFileName = App.Path & "\Reports\OPDocInvoice.rpt"
DocInvoice.DiscardSavedData = True
DocInvoice.ReplaceSelectionFormula ("{Appointment_Bill_Payment.Appointment_Bill_ID} = '" & strTXT & "'")

DocInvoice.WindowState = crptMaximized
DocInvoice.Action = 1


End Sub

Private Sub cmdsave_click()
Dim str As String
Dim BillPayID As String
Dim RowNo As Integer
Dim rsAddBill As New Recordset


RowNo = 1

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
      
    If optCash.Value = True Then
        cnPatients.Execute ("Insert into Appointment_Bill_Payment values('" & BillPayID & "','" & txtBillNo & "'," & Val(txtPayingAmt.Text) & ",'" & Format(dtpPayDate.Value, "mm/dd/yy") & "','" & str & "',Null,Null,Null)")
        MFG.TextMatrix(RowNo, 0) = RowNo
        MFG.TextMatrix(RowNo, 1) = Val(txtPayingAmt.Text)
        MFG.TextMatrix(RowNo, 2) = Format(dtpPayDate.Value, "dd-MMM-yyyy")
        MFG.TextMatrix(RowNo, 3) = str
        RowNo = RowNo + 1
        MFG.Rows = MFG.Rows + 1
        cmdSave.Enabled = False
        cmdPrint.Enabled = True
    Else
        If txtDDNo.Text = "" Or cmbBank.Text = "" Then
            MsgBox "DD Number or Bank Name Not Found...", vbCritical + vbOKOnly
            txtDDNo.SetFocus
            Exit Sub
        End If
        con.Execute ("Insert into Appointment_Bill_Payment values(" & BillPayID & "," & txtBillNo & "," & Val(txtPayingAmt.Text) & ",'" & Format(dtpPayDate.Value, "mm/dd/yy") & "','" & str & "','" & txtDDNo.Text & "','" & Format(dtpDDDate.Value, "mm/dd/yy") & "','" & cmbBank.Text & "')")
        MFG.TextMatrix(RowNo, 0) = RowNo
        MFG.TextMatrix(RowNo, 1) = Val(txtPayingAmt.Text)
        MFG.TextMatrix(RowNo, 2) = Format(dtpPayDate.Value, "dd-MMM-yyyy")
        MFG.TextMatrix(RowNo, 3) = str
        MFG.TextMatrix(RowNo, 4) = txtDDNo.Text
        MFG.TextMatrix(RowNo, 5) = Format(dtpDDDate.Value, "dd-MMM-yyyy")
        MFG.TextMatrix(RowNo, 6) = cmbBank.Text
        RowNo = RowNo + 1
        MFG.Rows = MFG.Rows + 1
        cmdSave.Enabled = False
        cmdPrint.Enabled = True
    End If
  


End If



End Sub

Private Sub Form_Load()
Call Functions.DisableMenu
Me.WindowState = vbMaximized
Call setFormData
Call Refresh_Data
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 22 Then KeyAscii = 0: Exit Sub
    KeyAscii = DataEntryValidation(KeyAscii, ActiveControl.Tag)
End Sub


Private Sub Refresh_Data()


Dim i As Integer
i = 0

MFG.clear
MFG.ColWidth(0) = 1000
MFG.ColAlignment(0) = 4
For i = 1 To 6 Step 1
    MFG.ColWidth(i) = 2000
    MFG.ColAlignment(i) = 4
Next i
MFG.TextMatrix(0, 0) = "REC NO"
MFG.TextMatrix(0, 1) = "AMOUNT PAID"
MFG.TextMatrix(0, 2) = "PAID DATE"
MFG.TextMatrix(0, 3) = "PAY TYPE"
MFG.TextMatrix(0, 4) = "DD/CHEQUE NO"
MFG.TextMatrix(0, 5) = "DD DATE"
MFG.TextMatrix(0, 6) = "BANK"





End Sub

Private Sub Label12_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Functions.EnableMenu
End Sub

Private Sub optCash_Click()
If optCash.Value = True Then
    txtDDNo.Text = ""
    txtDDNo.Locked = True
    cmbBank.Locked = True
    dtpDDDate.Enabled = False
Else
    txtDDNo.Locked = False
    cmbBank.Locked = False
    dtpDDDate.Enabled = True
End If
End Sub

Private Sub optCheque_Click()
If optCheque.Value = True Then
    txtDDNo.Locked = False
    cmbBank.Locked = False
    dtpDDDate.Enabled = True
Else
    txtDDNo.Text = ""
    txtDDNo.Locked = True
    cmbBank.Locked = True
    dtpDDDate.Enabled = False
End If
End Sub

Private Sub optDD_Click()
If optDD.Value = True Then
    txtDDNo.Locked = False
    cmbBank.Locked = False
    dtpDDDate.Enabled = True
Else
    txtDDNo.Text = ""
    txtDDNo.Locked = True
    cmbBank.Locked = True
    dtpDDDate.Enabled = False
End If
End Sub

Private Sub optOthers_Click()
If optOthers.Value = True Then
    txtDDNo.Locked = False
    cmbBank.Locked = False
    dtpDDDate.Enabled = True
Else
    txtDDNo.Text = ""
    txtDDNo.Locked = True
    cmbBank.Locked = True
    dtpDDDate.Enabled = False
End If
End Sub

Private Sub txtCustomerAdv_Change()

End Sub

Private Sub txtPaidAmt_Change()

End Sub

Private Sub txtPayingAmt_KeyPress(KeyAscii As Integer)
Dim str As String
str = "0123456789."
If KeyAscii = 13 And txtPayingAmt.Text <> "" Then
    cmdSave.SetFocus
End If
If InStr(str, Chr(KeyAscii)) = 0 And KeyAscii <> 8 Then
    KeyAscii = 0
End If
End Sub

Private Sub txtPayingAmt_LostFocus()
If txtPayingAmt.Text <> "" Then
    If Val(txtPayingAmt.Text) = 0 Then
        MsgBox "Paying Amount Cannot Be Zero...", vbInformation + vbOKOnly
        txtPayingAmt.SetFocus
        Exit Sub
    End If
    If Val(txtPayingAmt.Text) > Val(txtNetValue.Text) Then
        MsgBox "Paying Amount Cannot Be Greater Than Total Amount...", vbCritical + vbOKOnly
        txtPayingAmt.Text = ""
        txtPayingAmt.SetFocus
        Exit Sub
    End If
     txtBalAmt.Text = Round((Val(txtNetValue.Text) - Val(txtPayingAmt.Text)), 2)
     
    If Val(txtBalAmt.Text) = 0 Then
    txtBillStatus.Text = "Paid"
        Else
    txtBillStatus.Text = "Un-Paid"
    End If
End If
End Sub
