VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmSalesReport 
   BackColor       =   &H00FF8080&
   Caption         =   "Pharmacy Sales Report"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10500
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSalesReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   10500
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF8080&
      Caption         =   "Customer ID"
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
      Height          =   1575
      Left            =   1200
      TabIndex        =   4
      Top             =   6480
      Width           =   8175
      Begin VB.CommandButton cmdViewReport4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Report"
         Height          =   1095
         Left            =   6960
         Picture         =   "frmSalesReport.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtCustomerID 
         Height          =   375
         Left            =   3840
         TabIndex        =   16
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Customer ID"
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
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "Criteria"
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
      Height          =   1575
      Left            =   1200
      TabIndex        =   3
      Top             =   4560
      Width           =   8175
      Begin VB.CommandButton cmdViewReport3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Report"
         Height          =   1095
         Left            =   6960
         Picture         =   "frmSalesReport.frx":5D6E
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox cmbOrder 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtOrder 
         Height          =   375
         Left            =   3960
         TabIndex        =   13
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF8080&
         Caption         =   " = "
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
         Left            =   3240
         TabIndex        =   15
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Amount"
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
      Left            =   1200
      TabIndex        =   2
      Top             =   2760
      Width           =   8175
      Begin VB.CommandButton cmdViewReport2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Report"
         Height          =   1095
         Left            =   6960
         Picture         =   "frmSalesReport.frx":62FA
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtTotalFrom 
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Tag             =   "Amt"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtTotalto 
         Height          =   375
         Left            =   5040
         TabIndex        =   9
         Tag             =   "Amt"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF8080&
         Caption         =   "Total Amount"
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
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF8080&
         Caption         =   "To"
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
         Left            =   3960
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Range"
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
      Left            =   1200
      TabIndex        =   1
      Top             =   960
      Width           =   8175
      Begin VB.CommandButton cmdViewReport1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Report"
         Height          =   1095
         Left            =   6960
         Picture         =   "frmSalesReport.frx":6886
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPTo 
         Height          =   495
         Left            =   5040
         TabIndex        =   5
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         Format          =   20643841
         CurrentDate     =   38367
      End
      Begin MSComCtl2.DTPicker DTPFrom 
         Height          =   495
         Left            =   1920
         TabIndex        =   6
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393216
         Format          =   20643841
         CurrentDate     =   38367
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "From"
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
         TabIndex        =   8
         Top             =   480
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "To"
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
         Left            =   3840
         TabIndex        =   7
         Top             =   480
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   1095
      Left            =   8160
      Picture         =   "frmSalesReport.frx":6E12
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8160
      Width           =   1095
   End
   Begin Crystal.CrystalReport crSale 
      Left            =   720
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowSearchBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      FormHeightDT    =   9930
      FormWidthDT     =   10620
      FormScaleHeightDT=   9420
      FormScaleWidthDT=   10500
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "PHARMACY SALES REPORT"
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
      Left            =   2640
      TabIndex        =   22
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "frmSalesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdViewReport1_Click()
'Dim strReport As String
'strReport = App.Path & "\Reports\Pharmacy\Sales.rpt"


crSale.ReportFileName = App.Path & "\Reports\Pharmacy\Sales.rpt"
crSale.DiscardSavedData = True
crSale.ReplaceSelectionFormula ("{Orders.OrderDate}   >=#" & DTPFrom & "#  and {Orders.OrderDate}  <=#" & DTPTo & "#")


crSale.WindowState = crptMaximized
crSale.Action = 1
End Sub

Private Sub cmdViewReport3_Click()

crSale.ReportFileName = App.Path & "\Reports\Pharmacy\Sales.rpt"
crSale.DiscardSavedData = True
If cmbOrder.ListIndex = 3 Or cmbOrder.ListIndex = 4 Or cmbOrder.ListIndex = 5 Or cmbOrder.ListIndex = 6 Then
    crSale.ReplaceSelectionFormula ("{OrderDetails." & cmbOrder & "}  =" & txtOrder & "")
Else
    crSale.ReplaceSelectionFormula ("{OrderDetails." & cmbOrder & "}  ='" & txtOrder & "'")
End If

crSale.WindowState = crptMaximized
crSale.Action = 1
End Sub

Private Sub cmdViewReport4_Click()
Dim strReport As String
strReport = App.Path & "\Reports\Pharmacy\Sales.rpt"


crSale.ReportFileName = App.Path & "\Reports\Pharmacy\Sales.rpt"
crSale.DiscardSavedData = True
crSale.ReplaceSelectionFormula ("{Orders.CustomerID}  ='" & txtCustomerID & "'")


crSale.WindowState = crptMaximized
crSale.Action = 1


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 22 Then KeyAscii = 0: Exit Sub
    KeyAscii = DataEntryValidation(KeyAscii, ActiveControl.Tag)
End Sub
Private Sub cmdViewReport2_Click()


Dim strReport As String
strReport = App.Path & "\Reports\Pharmacy\Sales.rpt"


crSale.ReportFileName = App.Path & "\Reports\Pharmacy\Sales.rpt"
crSale.DiscardSavedData = True
crSale.ReplaceSelectionFormula ("{OrderDetails.NetValue}   >=" & Val(txtTotalFrom) & "  and {OrderDetails.NetValue}  <=" & Val(txtTotalto) & "")


crSale.WindowState = crptMaximized
crSale.Action = 1
End Sub

Private Sub Form_Load()
Dim rsadd As Recordset
Set rsadd = New ADODB.Recordset

rsadd.Open "Select * from OrderDetails", cnPatients, adOpenDynamic, adLockReadOnly

For i = 0 To rsadd.Fields.Count - 1 Step 1
    cmbOrder.AddItem rsadd.Fields(i).name, i
Next
rsadd.Close
End Sub
