VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmCusSalesReport 
   Caption         =   "Sales By Customer"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   7500
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport crCusSale 
      Left            =   2400
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   6120
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   2895
   End
   Begin VB.ComboBox cmbCustomer 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "frmCusSalesReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim rsAddCus As Recordset
Set rsAddCus = New ADODB.Recordset

rsAddCus.Open "select * from Customers", cnPatients, adOpenDynamic, adLockReadOnly

For i = 0 To rsAddCus.Fields.Count - 1 Step 1
    


End Sub
