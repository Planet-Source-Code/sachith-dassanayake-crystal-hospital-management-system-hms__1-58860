VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmOrder 
   BackColor       =   &H00FF8080&
   Caption         =   "Invoice"
   ClientHeight    =   11250
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   12345
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11250
   ScaleWidth      =   12345
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Remove Selected Item"
      Height          =   375
      Left            =   9960
      TabIndex        =   41
      Top             =   8760
      Width           =   2175
   End
   Begin VB.CommandButton cndew 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&New"
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
      Left            =   3720
      Picture         =   "frmOrder.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   10080
      Width           =   1185
   End
   Begin Crystal.CrystalReport PInvoice 
      Left            =   840
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton cmdAddList 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Add to the List"
      Default         =   -1  'True
      Height          =   975
      Left            =   10680
      Picture         =   "frmOrder.frx":5C69
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5160
      Width           =   1335
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
      Left            =   5160
      Picture         =   "frmOrder.frx":6104
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Click To Save Bill Payment Information"
      Top             =   10080
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
      Left            =   8040
      Picture         =   "frmOrder.frx":65B2
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Click To Close"
      Top             =   10080
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Print Invoice"
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
      Height          =   915
      Left            =   6600
      Picture         =   "frmOrder.frx":6AB6
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   10080
      Width           =   1215
   End
   Begin VB.CommandButton cmdVProducts 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4440
      TabIndex        =   33
      Top             =   3240
      Width           =   495
   End
   Begin VB.ComboBox cmbPID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   3240
      Width           =   1935
   End
   Begin VB.ComboBox cmbCID 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2280
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtpayable 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   9300
      Width           =   1455
   End
   Begin VB.TextBox txtdisgvn 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   9300
      Width           =   1575
   End
   Begin VB.TextBox txtgrndtot 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   9300
      Width           =   1095
   End
   Begin VB.TextBox txtBillID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Medicine Details"
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
      Height          =   2775
      Left            =   480
      TabIndex        =   1
      Top             =   2280
      Width           =   11535
      Begin VB.TextBox txtStock 
         Height          =   285
         Left            =   6480
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   1560
         Width           =   1215
      End
      Begin VB.ComboBox cmbPtID 
         Height          =   315
         Left            =   6480
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox txtmedname 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txttotamt 
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox txtdis 
         Height          =   285
         Left            =   10320
         TabIndex        =   5
         Tag             =   "Amt"
         Top             =   930
         Width           =   975
      End
      Begin VB.TextBox txtAmount 
         Enabled         =   0   'False
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtrpu 
         Height          =   285
         Left            =   10320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtqty 
         Height          =   285
         Left            =   6480
         TabIndex        =   2
         Tag             =   "Num"
         Top             =   2160
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPIssue 
         Height          =   375
         Left            =   6480
         TabIndex        =   27
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
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
         Format          =   45809665
         CurrentDate     =   38353
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "Units Available"
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
         Left            =   4800
         TabIndex        =   32
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C96C59&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Code"
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
         TabIndex        =   29
         Top             =   375
         Width           =   1470
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category ID"
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
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Height          =   195
         Left            =   8760
         TabIndex        =   15
         Top             =   1560
         Width           =   750
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Given"
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
         Left            =   8760
         TabIndex        =   14
         Top             =   930
         Width           =   1455
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Total"
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
         Left            =   8760
         TabIndex        =   13
         Top             =   2160
         Width           =   1020
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Per Unit"
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
         Left            =   8760
         TabIndex        =   12
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         TabIndex        =   11
         Top             =   2160
         Width           =   825
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issue Date"
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
         TabIndex        =   10
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine Name"
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
         Left            =   120
         TabIndex        =   9
         Top             =   1560
         Width           =   1470
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product ID"
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
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1035
      End
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   375
      Left            =   10200
      TabIndex        =   0
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   45809665
      CurrentDate     =   38353
   End
   Begin MSFlexGridLib.MSFlexGrid MFG 
      Height          =   2415
      Left            =   240
      TabIndex        =   21
      Top             =   6240
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   4260
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      ForeColorSel    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      FormHeightDT    =   11760
      FormWidthDT     =   12465
      FormScaleHeightDT=   11250
      FormScaleWidthDT=   12345
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "PHARMACY SALES INVOICE"
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
      Left            =   3397
      TabIndex        =   38
      Top             =   360
      Width           =   5550
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   615
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   11535
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount Payable"
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
      Left            =   7560
      TabIndex        =   26
      Top             =   9375
      Width           =   1980
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount Given"
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
      Left            =   4080
      TabIndex        =   25
      Top             =   9360
      Width           =   1455
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
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
      TabIndex        =   24
      Top             =   9375
      Width           =   1140
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
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
      Left            =   9120
      TabIndex        =   23
      Top             =   1560
      Width           =   810
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No"
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
      TabIndex        =   22
      Top             =   1440
      Width           =   615
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim stock As Integer
Dim OID As String
Private Sub cmbCID_Click()

Dim rsProducts As Recordset
Set rsProducts = New ADODB.Recordset
cmbPID.clear

rsProducts.Open "Select * from Medicine_Details where CategoryID = '" & cmbCID & "' ", cnPatients, adOpenDynamic, adLockOptimistic

If rsProducts.EOF = False Then
    rsProducts.MoveFirst

    While rsProducts.EOF = False
        cmbPID.AddItem rsProducts(0)
        cmbPID.Text = rsProducts(0)
        rsProducts.MoveNext
    Wend
End If
If cmbPID.ListCount = 0 Then
txtmedname = ""
txtRPU = "0"
End If

rsProducts.Close

End Sub

Private Sub cmbPID_Click()

Dim rsProductName As Recordset
Set rsProductName = New ADODB.Recordset


rsProductName.Open "Select * from Medicine_Details where ProductID = '" & cmbPID & "'", cnPatients, adOpenDynamic, adLockReadOnly


If rsProductName.RecordCount > 1 Then
    MsgBox " Database Error"
    stock = 0
    txtStock = stock
    Exit Sub
ElseIf rsProductName.RecordCount = 0 Then
    txtmedname = ""
    txtRPU = "0.00"
    stock = 0

Else
    txtmedname = rsProductName(1)
    txtRPU = rsProductName(5)
    stock = rsProductName(4)
    
End If

txtStock = stock

rsProductName.Close

End Sub



Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim selectedRow As Integer
selectedRow = MFG.Row


If selectedRow = MFG.Rows - 1 Then
    MsgBox "Invalid Selection.", vbCritical
    Exit Sub
End If


If Not MFG.TextMatrix(1, 1) = "" Then
    MFG.RemoveItem (selectedRow)
    Call CalFinal
End If



End Sub

Private Sub cmdVProducts_Click()
FrmVProducts.Show
End Sub

Private Sub cndew_Click()
Dim ctl As Control

For Each ctl In Controls
    If TypeOf ctl Is TextBox Then
        ctl.Text = ""
    End If
Next
cmbPtID.Enabled = True
cmdSave.Enabled = True
MFG.clear
MFG.Refresh
MFG.Rows = 2
Call Form_Load
End Sub

Private Sub Command1_Click()
Dim strReport As String
Dim strTXT As String

strTXT = txtBillID.Text
strReport = App.Path & "\Reports\Pharmacy\invoice.rpt"
PInvoice.DiscardSavedData = True
PInvoice.ReportFileName = strReport
PInvoice.ReplaceSelectionFormula ("{OrderDetails.OrderID} = '" & strTXT & "'")
PInvoice.WindowState = crptMaximized
PInvoice.Action = 1
End Sub

Private Sub Command2_Click()
MsgBox MFG.Row
MsgBox MFG.Rows
End Sub

Private Sub Form_Load()

Call SetData
Call BillID
Call CustDetails
Call MFGVALUES
Command1.Enabled = False
DTPDate.Value = Date
DTPIssue.Value = Date

End Sub

Public Sub SetData()

Dim rsCategories As Recordset
Set rsCategories = New ADODB.Recordset

rsCategories.Open "select * from Medicine_Categories", cnPatients, adOpenDynamic, adLockOptimistic
cmbCID.clear
While rsCategories.EOF = False
cmbCID.AddItem rsCategories(0)
rsCategories.MoveNext

Wend
rsCategories.Close


End Sub

Public Sub BillID()


   Dim BID As String
   Dim rsOrderID As Recordset
    Set rsOrderID = New ADODB.Recordset
    ' Generatin Order Details ID
        
    BID = Functions.UID(6, "MODRID_")
    rsOrderID.Open " Select * from Orders", cnPatients, adOpenDynamic, adLockPessimistic
    If rsOrderID.EOF = False Then
    While rsOrderID.EOF = False
        If rsOrderID(0) = BID Then
            BID = Functions.UID(6, "MODRID_")
            rsOrderID.MoveFirst
        End If
    rsOrderID.MoveNext
    Wend
    End If



txtBillID = BID

End Sub

Public Sub MFGVALUES()
MFG.TextMatrix(0, 1) = "ORDER ID"
MFG.TextMatrix(0, 2) = "PRODUCT ID"
MFG.TextMatrix(0, 3) = "PRODUCT NAME"
MFG.TextMatrix(0, 4) = "QUANTITY"
MFG.TextMatrix(0, 5) = "UNIT PRICE"
MFG.TextMatrix(0, 6) = "DISCOUNT"
MFG.TextMatrix(0, 7) = "TOTAL AMOUNT"
Functions.SizeColumnHeaders MFG, Me

End Sub

Public Sub CustDetails()

Dim rsAddCust As Recordset
Set rsAddCust = New ADODB.Recordset

rsAddCust.Open "Select * from Customers", cnPatients, adOpenDynamic, adLockReadOnly

cmbPtID.clear
If rsAddCust.EOF = False Then
rsAddCust.MoveFirst

While rsAddCust.EOF = False
    cmbPtID.AddItem rsAddCust(0)
    rsAddCust.MoveNext
Wend


End If

rsAddCust.Close


End Sub







Private Sub txtdis_Change()
txtAmount = Val(txtRPU) * Val(txtqty)
txttotamt = Val(txtAmount) - Val(txtdis)
End Sub

Private Sub txtqty_Change()

txtAmount = Val(txtRPU) * Val(txtqty)
txttotamt = Val(txtAmount) - Val(txtdis)
End Sub


Private Sub cmdAddList_Click()



'On Error Resume Next
    Dim rsMed As Recordset
    Dim i As Integer
    
   
    
    
If MFG.Rows > 2 Then
    For i = 1 To MFG.Rows - 2 Step 1
        If MFG.TextMatrix(i, 2) = cmbPID Then
            MsgBox "Medicine Already Exist In The List Cannot Add Same Medicine Again.....", vbCritical + vbOKOnly
            Exit Sub
        End If
    Next i
End If




If txtAmount = "" Or txttotamt = "" Or txtqty = "" Or txtRPU = "" Then
    MsgBox "Please Enter the relevant Fields"
    Exit Sub
End If
If Val(txtqty) = 0 Then
    MsgBox "Quantity Cannot be Zero", vbCritical
    Exit Sub
End If
If Val(txtqty) > stock Then
    MsgBox "The Quantity Cannot be greater than Stock", vbCritical
    Exit Sub
End If
If cmbPtID = "" Then
    MsgBox "Please Select the Customer ID", vbCritical
    cmbPtID.SetFocus
    Exit Sub
End If


Row = MFG.Rows - 1
With MFG

        .Rows = .Rows + 1
                
        MFG.TextMatrix(Row, 1) = txtBillID
        MFG.TextMatrix(Row, 2) = cmbPID
        MFG.TextMatrix(Row, 3) = txtmedname
        MFG.TextMatrix(Row, 4) = txtqty
        MFG.TextMatrix(Row, 5) = txtRPU
        MFG.TextMatrix(Row, 6) = txtdis
        MFG.TextMatrix(Row, 7) = txttotamt
        
  
    .FixedRows = 1
    .RowHeight(0) = .RowHeight(1) * 1.5
     Functions.SizeColumns MFG, Me
     MFGVALUES
     
     Row = Row + 1
     
End With

cmbPtID.Enabled = False

Call CalFinal
Call TextClear

End Sub
Public Sub CalFinal()
Dim amount As Double
Dim Discount As Double
Dim Total As Double

If MFG.Rows > 2 Then
    For i = 1 To MFG.Rows - 2 Step 1
        amount = amount + (Val(MFG.TextMatrix(i, 5)) * Val(MFG.TextMatrix(i, 4)))
        Discount = Discount + Val(MFG.TextMatrix(i, 6))
        Total = Total + Val(MFG.TextMatrix(i, 7))
        
    Next i
End If
txtgrndtot = amount
txtdisgvn = Discount
txtpayable = Total
Debug.Print Val(amount) - Val(Discount)
End Sub

Public Sub TextClear()
txtqty = ""
txtdis = ""
txtAmount = ""
txttotamt = ""

cmbPID_Click
End Sub



Private Sub cmdSave_Click()
    Dim rsOrderID As Recordset

   If MFG.Rows = 2 Then
    MsgBox "Please Add Items to list before you save", vbCritical, "Error Occured"
    Exit Sub
   End If
   
Dim flag, flag1, flag2 As Boolean
flag = False
flag1 = False
flag2 = False

   
    Set rsOrderID = New ADODB.Recordset
   
    rsOrderID.Open " Select * from Orders", cnPatients, adOpenDynamic, adLockPessimistic
      
    
    rsOrderID.AddNew
        rsOrderID(0) = txtBillID
        rsOrderID(1) = cmbPtID
        rsOrderID(2) = DTPDate
    rsOrderID.Update
    flag2 = True
    
    rsOrderID.Close
    
Dim rsMed As Recordset
Set rsMed = New ADODB.Recordset
Dim MID As String
Dim RQuantity As Integer

Dim rsAddPatient As Recordset
Set rsAddPatient = New ADODB.Recordset
Dim rsStock As Recordset
Set rsStock = New ADODB.Recordset

    

rsMed.Open "SELECT * FROM OrderDetails", cnPatients, adOpenDynamic, adLockPessimistic

For i = 1 To MFG.Rows - 2 Step 1

    ' Generatin Order Details ID
    MID = Functions.UID(6, "ODRDTL_")
    rsAddPatient.Open " Select * from OrderDetails", cnPatients, adOpenDynamic, adLockReadOnly
    If rsAddPatient.EOF = False Then
    While rsAddPatient.EOF = False
        If rsAddPatient(0) = MID Then
            MID = Functions.UID(6, "ODRDTL_")
            rsAddPatient.MoveFirst
        End If
    rsAddPatient.MoveNext
    Wend
    End If
    rsAddPatient.Close

        With rsMed
            .AddNew
                !OrderDetailID = MID
                !OrderID = txtBillID
                !ProductID = MFG.TextMatrix(i, 2)
                !QUANTITY = MFG.TextMatrix(i, 4)
                !UNITPRICE = MFG.TextMatrix(i, 5)
                !Discount = Val(MFG.TextMatrix(i, 6))
                !NetValue = (Val(MFG.TextMatrix(i, 4)) * Val(MFG.TextMatrix(i, 5))) - Val(MFG.TextMatrix(i, 6))
            .Update
            flag = True
        End With
        
        ' Substract the stock from the products
        rsStock.Open "select * from Medicine_Details where ProductID= '" & MFG.TextMatrix(i, 2) & "'", cnPatients, adOpenDynamic, adLockPessimistic
        If rsStock.EOF = False Then
            rsStock(4) = rsStock(4) - Val(MFG.TextMatrix(i, 4))
            rsStock.Update
            flag1 = True
        End If
        rsStock.Close
Next

rsMed.Close

If flag = True And flag1 = True And flag2 = True Then
    MsgBox "Record Saved Succesfully !!"
Else
    MsgBox "Error Updating Record", vbCritical
End If

Command1.Enabled = True
cmdSave.Enabled = False



End Sub

Private Sub txtqty_LostFocus()
If Val(txtqty) > stock Then
    MsgBox "The Quantity Cannot be greater than Stock", vbCritical
End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 22 Then KeyAscii = 0: Exit Sub
    KeyAscii = DataEntryValidation(KeyAscii, ActiveControl.Tag)
End Sub

