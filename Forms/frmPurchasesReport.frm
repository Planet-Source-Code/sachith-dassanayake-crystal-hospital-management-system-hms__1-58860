VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmPurchasesReport 
   BackColor       =   &H00FF8080&
   Caption         =   "Pharmacy Purchases Report"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9465
   ScaleWidth      =   8850
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   1095
      Left            =   7320
      Picture         =   "frmPurchasesReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8160
      Width           =   1095
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
      Left            =   360
      TabIndex        =   16
      Top             =   960
      Width           =   8175
      Begin VB.CommandButton cmdViewReport1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Report"
         Height          =   1095
         Left            =   6960
         Picture         =   "frmPurchasesReport.frx":0504
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPTo 
         Height          =   495
         Left            =   5040
         TabIndex        =   18
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         _Version        =   393216
         Format          =   46071809
         CurrentDate     =   38367
      End
      Begin MSComCtl2.DTPicker DTPFrom 
         Height          =   495
         Left            =   1920
         TabIndex        =   19
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393216
         Format          =   46071809
         CurrentDate     =   38367
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
         TabIndex        =   21
         Top             =   480
         Width           =   240
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
         TabIndex        =   20
         Top             =   480
         Width           =   510
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
      Left            =   360
      TabIndex        =   10
      Top             =   2760
      Width           =   8175
      Begin VB.TextBox txtTotalto 
         Height          =   375
         Left            =   5040
         TabIndex        =   13
         Tag             =   "Amt"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtTotalFrom 
         Height          =   375
         Left            =   1920
         TabIndex        =   12
         Tag             =   "Amt"
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdViewReport2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Report"
         Height          =   1095
         Left            =   6960
         Picture         =   "frmPurchasesReport.frx":0A90
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1095
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
         TabIndex        =   15
         Top             =   600
         Width           =   855
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
         TabIndex        =   14
         Top             =   480
         Width           =   1575
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
      Left            =   360
      TabIndex        =   5
      Top             =   4560
      Width           =   8175
      Begin VB.TextBox txtOrder 
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   480
         Width           =   2775
      End
      Begin VB.ComboBox cmbOrder 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   2295
      End
      Begin VB.CommandButton cmdViewReport3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Report"
         Height          =   1095
         Left            =   6960
         Picture         =   "frmPurchasesReport.frx":101C
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   1095
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
         TabIndex        =   9
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF8080&
      Caption         =   "Supplier ID"
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
      Left            =   360
      TabIndex        =   1
      Top             =   6480
      Width           =   8175
      Begin VB.TextBox txtCustomerID 
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdViewReport4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Report"
         Height          =   1095
         Left            =   6960
         Picture         =   "frmPurchasesReport.frx":15A8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FF8080&
         Caption         =   "Supplier ID"
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
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport crPurchase 
      Left            =   840
      Top             =   360
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
      Left            =   120
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      FormHeightDT    =   9975
      FormWidthDT     =   8970
      FormScaleHeightDT=   9465
      FormScaleWidthDT=   8850
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "PURCHASE REPORTS"
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
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   4170
   End
End
Attribute VB_Name = "frmPurchasesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdViewReport1_Click()
Dim strReport As String
strReport = App.Path & "\Reports\Pharmacy\Purchase.rpt"


crPurchase.ReportFileName = App.Path & "\Reports\Pharmacy\Purchase.rpt"
crPurchase.DiscardSavedData = True
crPurchase.ReplaceSelectionFormula ("{Purchase_Orders.PurchaseOrderDate}   >=#" & DTPFrom & "#  and {Purchase_Orders.PurchaseOrderDate}  <=#" & DTPTo & "#  ")

crPurchase.WindowState = crptMaximized
crPurchase.Action = 1
End Sub

Private Sub cmdViewReport2_Click()
Dim strReport As String
strReport = App.Path & "\Reports\Pharmacy\Purchase.rpt"

crPurchase.ReportFileName = App.Path & "\Reports\Pharmacy\Purchase.rpt"
crPurchase.DiscardSavedData = True
crPurchase.ReplaceSelectionFormula ("{Purchase_Orde_Details.NetValue}   >=" & Val(txtTotalFrom) & "  and {Purchase_Orde_Details.NetValue}  <=" & Val(txtTotalto) & "")


crPurchase.WindowState = crptMaximized
crPurchase.Action = 1
End Sub

Private Sub cmdViewReport3_Click()
Dim strReport As String
strReport = App.Path & "\Reports\Pharmacy\Purchase.rpt"


crPurchase.ReportFileName = App.Path & "\Reports\Pharmacy\Purchase.rpt"
crPurchase.DiscardSavedData = True

If cmbOrder.ListIndex = 3 Or cmbOrder.ListIndex = 4 Or cmbOrder.ListIndex = 5 Or cmbOrder.ListIndex = 6 Then
crPurchase.ReplaceSelectionFormula ("{Purchase_Orde_Details." & cmbOrder & "}  =" & txtOrder & "")
Else
crPurchase.ReplaceSelectionFormula ("{Purchase_Orde_Details." & cmbOrder & "}  ='" & txtOrder & "'")
End If

crPurchase.WindowState = crptMaximized
crPurchase.Action = 1
End Sub

Private Sub cmdViewReport4_Click()
Dim strReport As String
strReport = App.Path & "\Reports\Pharmacy\Purchase.rpt"


crPurchase.ReportFileName = App.Path & "\Reports\Pharmacy\Purchase.rpt"
crPurchase.DiscardSavedData = True

crPurchase.ReplaceSelectionFormula ("{Purchase_Orders.PurchaseSupplierID}  ='" & txtCustomerID & "'")


crPurchase.WindowState = crptMaximized
crPurchase.Action = 1

End Sub

Private Sub Form_Load()
Dim rsadd As Recordset
Set rsadd = New ADODB.Recordset

rsadd.Open "Select * from Purchase_Orde_Details", cnPatients, adOpenDynamic, adLockReadOnly

For i = 0 To rsadd.Fields.Count - 1 Step 1
    cmbOrder.AddItem rsadd.Fields(i).name, i
Next
rsadd.Close
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 22 Then KeyAscii = 0: Exit Sub
    KeyAscii = DataEntryValidation(KeyAscii, ActiveControl.Tag)
End Sub


