VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form FrmPurchases 
   BackColor       =   &H00FF8080&
   Caption         =   "Purchases"
   ClientHeight    =   9990
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   13005
   Icon            =   "FrmPurchases.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9990
   ScaleWidth      =   13005
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cndew 
      BackColor       =   &H00FFFFFF&
      Caption         =   "New"
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
      Left            =   4200
      Picture         =   "FrmPurchases.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   9000
      Width           =   1185
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Remove Selected Item"
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
      Left            =   10560
      TabIndex        =   40
      Top             =   7920
      Width           =   2175
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
      Left            =   7080
      Picture         =   "FrmPurchases.frx":5C69
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Click To Close"
      Top             =   9000
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
      Left            =   5640
      Picture         =   "FrmPurchases.frx":616D
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Click To Save Bill Payment Information"
      Top             =   9000
      Width           =   1185
   End
   Begin VB.CommandButton cmdAddList 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Add to the List"
      Default         =   -1  'True
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
      Left            =   11520
      Picture         =   "FrmPurchases.frx":661B
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   4920
      Width           =   1215
   End
   Begin VB.TextBox txtpayable 
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
      Left            =   10560
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   8400
      Width           =   1455
   End
   Begin VB.TextBox txtdisgvn 
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
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   8400
      Width           =   1575
   End
   Begin VB.TextBox txtgrndtot 
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
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   8400
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MFG 
      Height          =   1815
      Left            =   480
      TabIndex        =   27
      Top             =   6000
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   8
      SelectionMode   =   1
   End
   Begin VB.TextBox txtBillID 
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Height          =   2535
      Left            =   360
      TabIndex        =   0
      Top             =   2280
      Width           =   12375
      Begin VB.CommandButton cmdVSuppliers 
         Caption         =   "..."
         Height          =   255
         Left            =   4080
         TabIndex        =   38
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdVProducts 
         Caption         =   "..."
         Height          =   255
         Left            =   8520
         TabIndex        =   37
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtNet 
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
         Left            =   10800
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtAmount 
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
         Left            =   10800
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtdis 
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
         Left            =   10800
         TabIndex        =   22
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox txtRPU 
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
         Left            =   10800
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtUPurchased 
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
         Left            =   6840
         TabIndex        =   18
         Tag             =   "Num"
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtunits 
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
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtPName 
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
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   840
         Width           =   2175
      End
      Begin VB.ComboBox cmbPID 
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
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txtSCName 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox txtSName 
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
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   2535
      End
      Begin VB.ComboBox cmbSID 
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
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FF8080&
         Caption         =   "Net Amount"
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
         Left            =   9240
         TabIndex        =   24
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label11 
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
         Height          =   495
         Left            =   9240
         TabIndex        =   23
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FF8080&
         Caption         =   "Discount"
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
         Left            =   9240
         TabIndex        =   21
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FF8080&
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
         Height          =   495
         Left            =   9240
         TabIndex        =   19
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF8080&
         Caption         =   "Units Purchased"
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
         Left            =   5040
         TabIndex        =   17
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF8080&
         Caption         =   "Units In Stock"
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
         Left            =   5040
         TabIndex        =   15
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF8080&
         Caption         =   "Product Name"
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
         Left            =   5040
         TabIndex        =   12
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF8080&
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
         Height          =   255
         Left            =   5040
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF8080&
         Caption         =   "Contact Name"
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
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "Company Name"
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
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   375
      Left            =   10920
      TabIndex        =   8
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
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
      Format          =   45678593
      CurrentDate     =   38353
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      FormHeightDT    =   10500
      FormWidthDT     =   13125
      FormScaleHeightDT=   9990
      FormScaleWidthDT=   13005
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "PURCHASE ORDER"
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
      Left            =   4920
      TabIndex        =   39
      Top             =   480
      Width           =   3705
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
      Left            =   8160
      TabIndex        =   33
      Top             =   8520
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
      TabIndex        =   32
      Top             =   8460
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
      Left            =   960
      TabIndex        =   31
      Top             =   8475
      Width           =   1140
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
      Left            =   480
      TabIndex        =   10
      Top             =   1440
      Width           =   615
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
      Left            =   9840
      TabIndex        =   9
      Top             =   1560
      Width           =   810
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   12495
   End
End
Attribute VB_Name = "FrmPurchases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbPID_Click()

Dim rsProdName As Recordset
Set rsProdName = New ADODB.Recordset


rsProdName.Open "Select * from Medicine_Details where ProductID = '" & cmbPID & "'", cnPatients, adOpenDynamic, adLockReadOnly


If rsProdName.RecordCount > 1 Then
    MsgBox " Database Error"
    Exit Sub
Else
    txtPName = rsProdName(1)
    txtunits = rsProdName(4)
    txtRPU = rsProdName(5)
End If

rsProdName.Close

End Sub

Private Sub cmbSID_Click()

Dim rsSupplierName As Recordset
Set rsSupplierName = New ADODB.Recordset


rsSupplierName.Open "Select * from Suppliers where SupplierID = '" & cmbSID & "'", cnPatients, adOpenDynamic, adLockReadOnly


If rsSupplierName.RecordCount > 1 Then
    MsgBox " Database Error"
    Exit Sub
Else
    txtSName = rsSupplierName(1)
    txtSCName = rsSupplierName(2)
    
End If

rsSupplierName.Close


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

If txtAmount = "" Or txtNet = "" Or txtunits = "" Or txtRPU = "" Then
    MsgBox "Please Enter the relevant Fields"
    Exit Sub
End If
If Val(txtUPurchased) = 0 Then
    MsgBox "Quantity Cannot be Zero", vbCritical
    Exit Sub
End If




Row = MFG.Rows - 1
With MFG

        .Rows = .Rows + 1
                
        MFG.TextMatrix(Row, 1) = txtBillID
        MFG.TextMatrix(Row, 2) = cmbPID
        MFG.TextMatrix(Row, 3) = txtPName
        MFG.TextMatrix(Row, 4) = txtUPurchased
        MFG.TextMatrix(Row, 5) = txtRPU
        MFG.TextMatrix(Row, 6) = txtdis
        MFG.TextMatrix(Row, 7) = txtNet
        
          
  
    .FixedRows = 1
    .RowHeight(0) = .RowHeight(1) * 1.5
     Functions.SizeColumns MFG, Me
     MFGVALUES
     
     Row = Row + 1
     
End With


cmbSID.Enabled = False
Call CalcFinal
Call TextClear


End Sub
Public Sub CalcFinal()

Dim amount As Double
Dim Discount As Double
Dim Total As Double

If MFG.Rows > 2 Then
    For i = 1 To MFG.Rows - 2 Step 1
        amount = amount + (Val(MFG.TextMatrix(i, 5)) * Val(MFG.TextMatrix(i, 4)))
        Discount = Discount + MFG.TextMatrix(i, 6)
        Total = Total + MFG.TextMatrix(i, 7)
        
    Next i
End If
txtgrndtot = amount
txtdisgvn = Discount
txtpayable = Total
Debug.Print Val(amount) - Val(Discount)

End Sub

Public Sub TextClear()
txtunits = ""
txtUPurchased = ""
txtdis = ""
txtAmount = ""
txtNet = ""

cmbPID_Click
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
    Call CalcFinal
End If
End Sub

Private Sub cmdSave_Click()
   
   If MFG.Rows = 2 Then
    MsgBox "Please Add Items to list before you save", vbCritical, "Error Occured"
    Exit Sub
   End If

Dim flag, flag1, flag2 As Boolean

flag = False
flag1 = False
flag2 = False


    Dim rsOrderID As Recordset
    Dim OID As String
    Set rsOrderID = New ADODB.Recordset
    
    rsOrderID.Open " Select * from Purchase_Orders", cnPatients, adOpenDynamic, adLockPessimistic

    
    
    rsOrderID.AddNew
        rsOrderID(0) = txtBillID
        rsOrderID(1) = cmbSID
        rsOrderID(2) = DTPDate
        
    rsOrderID.Update
    flag = True
    
    rsOrderID.Close
    
Dim rsMed As Recordset
Set rsMed = New ADODB.Recordset
Dim MID As String
Dim RQuantity As Integer

Dim rsAddPatient As Recordset
Set rsAddPatient = New ADODB.Recordset
Dim rsStock As Recordset
Set rsStock = New ADODB.Recordset

    

rsMed.Open "SELECT * FROM Purchase_Orde_Details", cnPatients, adOpenDynamic, adLockPessimistic

For i = 1 To MFG.Rows - 2 Step 1

    ' Generating Purchase Order Details ID
    MID = Functions.UID(6, "PODRDTL_")
    rsAddPatient.Open " Select * from Purchase_Orde_Details", cnPatients, adOpenDynamic, adLockReadOnly
    If rsAddPatient.EOF = False Then
    While rsAddPatient.EOF = False
        If rsAddPatient(0) = MID Then
            MID = Functions.UID(6, "PODRDTL_")
            rsAddPatient.MoveFirst
        End If
    rsAddPatient.MoveNext
    Wend
    End If
    rsAddPatient.Close

        With rsMed
            .AddNew
                !PurchaseOrderDetailID = MID
                !PurchaseOrderID = txtBillID
                !PurchaseProductID = MFG.TextMatrix(i, 2)
                !PurchaseQUANTITY = Val(MFG.TextMatrix(i, 4))
                !PurchaseUnitPrice = Val(MFG.TextMatrix(i, 5))
                !PurchaseDiscount = Val(MFG.TextMatrix(i, 6))
                !NetValue = txtpayable
             .Update
             flag1 = True
        End With
        
        rsStock.Open "select * from Medicine_Details where ProductID= '" & MFG.TextMatrix(i, 2) & "'", cnPatients, adOpenDynamic, adLockPessimistic
        If rsStock.EOF = False Then
            rsStock(4) = rsStock(4) + Val(MFG.TextMatrix(i, 4))
            rsStock.Update
            flag2 = True
        End If
        rsStock.Close
Next
rsMed.Close
If flag = True And flag2 = True And flag1 = True Then
    MsgBox "Record Saved Succesfully !!", vbInformation, "Record Added"
    cmdSave.Enabled = False
Else
    MsgBox "An Error Occured while saving to the database", vbCritical
    Exit Sub
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
cmbSID.Enabled = True
cmdSave.Enabled = True
MFG.clear
MFG.Refresh
MFG.Rows = 2
Call Form_Load


End Sub

Private Sub Form_Load()

Call SetData
Call BillID
Call ProdDetails
Call MFGVALUES

End Sub

Public Sub SetData()

Dim rsSuppliers As Recordset
Set rsSuppliers = New ADODB.Recordset

  mbDataChanged = False
  rsSuppliers.Open "select * from Suppliers", cnPatients, adOpenDynamic, adLockOptimistic

rsSuppliers.MoveFirst

cmbSID.clear
While rsSuppliers.EOF = False
cmbSID.AddItem rsSuppliers(0)
rsSuppliers.MoveNext

Wend
rsSuppliers.Close


End Sub

Public Sub BillID()

   Dim BID As String
   Dim rsOrderID As Recordset
   Set rsOrderID = New ADODB.Recordset
    ' Generatin Order Details ID
        
    BID = Functions.UID(6, "MODRID_")
    rsOrderID.Open " Select * from Purchase_Orders", cnPatients, adOpenDynamic, adLockPessimistic
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
rsOrderID.Close

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

Public Sub ProdDetails()

Dim rsAddProd As Recordset
Set rsAddProd = New ADODB.Recordset

rsAddProd.Open "Select * from Medicine_Details", cnPatients, adOpenDynamic, adLockReadOnly


cmbPID.clear
If rsAddProd.EOF = False Then
rsAddProd.MoveFirst

While rsAddProd.EOF = False
    cmbPID.AddItem rsAddProd(0)
    cmbPID.Text = rsAddProd(0)
    rsAddProd.MoveNext
Wend


End If

rsAddProd.Close


End Sub



Private Sub txtdis_Change()
txtAmount = Val(txtRPU) * Val(txtUPurchased)
txtNet = Val(txtAmount) - Val(txtdis)
End Sub

Private Sub txtUPurchased_Change()

txtAmount = Val(txtRPU) * Val(txtUPurchased)
txtNet = Val(txtAmount) - Val(txtdis)
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 22 Then KeyAscii = 0: Exit Sub
    KeyAscii = DataEntryValidation(KeyAscii, ActiveControl.Tag)
End Sub

