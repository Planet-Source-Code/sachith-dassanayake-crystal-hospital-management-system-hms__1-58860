VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form FrmNavigation 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form Navigation"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   120
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      AutoCenterForm  =   -1  'True
      FormHeightDT    =   2670
      FormWidthDT     =   5520
      FormScaleHeightDT=   2190
      FormScaleWidthDT=   5430
   End
   Begin VB.CommandButton cmdPurchses 
      Caption         =   "Purchases"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdOrders 
      Caption         =   "Sale"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdSuppliers 
      Caption         =   "Suppliers"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdProducts 
      Caption         =   "Products"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdCustomers 
      Caption         =   "Customers"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton cmdCategories 
      Caption         =   "Categories"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   720
      Picture         =   "Navigation.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "FrmNavigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCategories_Click()
FrmCategories.Show
End Sub

Private Sub cmdCustomers_Click()
FrmCustomers.Show
End Sub

Private Sub cmdOrders_Click()
frmOrder.Show
End Sub

Private Sub cmdProducts_Click()
FrmProducts.Show
End Sub

Private Sub cmdPurchses_Click()
FrmPurchases.Show
End Sub

Private Sub cmdSuppliers_Click()
FrnSuppliers.Show
End Sub
