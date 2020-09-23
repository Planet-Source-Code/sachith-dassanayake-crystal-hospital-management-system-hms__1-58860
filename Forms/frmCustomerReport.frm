VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmCustomerReport 
   BackColor       =   &H00FF8080&
   Caption         =   "Customers Report"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   8190
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      Height          =   1095
      Left            =   4440
      Picture         =   "frmCustomerReport.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Frame Frame1 
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
      Height          =   2175
      Left            =   840
      TabIndex        =   2
      Top             =   1440
      Width           =   6855
      Begin VB.TextBox txtCustomer 
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   1200
         Width           =   2535
      End
      Begin VB.ComboBox cmbCustomer 
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF8080&
         Caption         =   "Item"
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
         Left            =   1080
         TabIndex        =   6
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "Option"
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
         Left            =   1080
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      FormHeightDT    =   5880
      FormWidthDT     =   8310
      FormScaleHeightDT=   5370
      FormScaleWidthDT=   8190
   End
   Begin Crystal.CrystalReport crCustomer 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton cmdViewReport 
      BackColor       =   &H00FFFFFF&
      Caption         =   "View Report"
      Height          =   1095
      Left            =   3120
      Picture         =   "frmCustomerReport.frx":0504
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "CUSTOMER BASED REPORT"
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
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   5415
   End
End
Attribute VB_Name = "frmCustomerReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdViewReport_Click()
Dim strReport As String
strReport = App.Path & "\Reports\Pharmacy\Customer Details.rpt"


crCustomer.ReportFileName = App.Path & "\Reports\Pharmacy\Customer Details.rpt"
crCustomer.DiscardSavedData = True
crCustomer.ReplaceSelectionFormula ("{Customers." & cmbCustomer & " }   ='" & txtCustomer & "'")


crCustomer.WindowState = crptMaximized
crCustomer.Action = 1
End Sub

Private Sub Form_Load()
Dim rsadd As Recordset
Set rsadd = New ADODB.Recordset

rsadd.Open "select * from Customers", cnPatients, adOpenDynamic, adLockReadOnly

For i = 0 To rsadd.Fields.Count - 1 Step 1
    cmbCustomer.AddItem rsadd(i).name, i
Next

rsadd.Close

End Sub

