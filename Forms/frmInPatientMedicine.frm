VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmInPatientMedicine 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   Caption         =   "In Patient Medicine Issue"
   ClientHeight    =   11145
   ClientLeft      =   2835
   ClientTop       =   1425
   ClientWidth     =   10020
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInPatientMedicine.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11145
   ScaleWidth      =   10020
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Remove Selected Item"
      Height          =   375
      Left            =   7680
      TabIndex        =   41
      Top             =   8640
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   285
      Left            =   4200
      TabIndex        =   40
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      Height          =   285
      Left            =   4200
      TabIndex        =   39
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtmedname 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   27
      Top             =   3840
      Width           =   2055
   End
   Begin VB.ComboBox cmbMedID 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   3240
      Width           =   2055
   End
   Begin VB.ComboBox cmbMedType 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox txttotamt 
      Height          =   285
      Left            =   7920
      TabIndex        =   24
      Tag             =   "Amt"
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox txtdis 
      Height          =   285
      Left            =   7920
      TabIndex        =   23
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox txtAmount 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7920
      TabIndex        =   22
      Tag             =   "Amt"
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox txtrpu 
      Height          =   285
      Left            =   2040
      TabIndex        =   21
      Tag             =   "Amt"
      Top             =   4800
      Width           =   2055
   End
   Begin VB.TextBox txtqty 
      Height          =   285
      Left            =   7920
      TabIndex        =   20
      Tag             =   "Num"
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtStock 
      Height          =   285
      Left            =   2040
      TabIndex        =   19
      Tag             =   "Num"
      Top             =   4320
      Width           =   2055
   End
   Begin MSFlexGridLib.MSFlexGrid MFG 
      Height          =   1695
      Left            =   240
      TabIndex        =   18
      Top             =   6840
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   2990
      _Version        =   393216
      Cols            =   9
      SelectionMode   =   1
      AllowUserResizing=   3
   End
   Begin MSComCtl2.DTPicker DTPDate 
      Height          =   285
      Left            =   7920
      TabIndex        =   17
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      Format          =   47448065
      CurrentDate     =   38353
   End
   Begin VB.TextBox txtName 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   8
      Top             =   2040
      Width           =   2055
   End
   Begin VB.ComboBox cmbAdmitID 
      Height          =   315
      Left            =   2040
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton cmdAddList 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Add to the List"
      Default         =   -1  'True
      Height          =   855
      Left            =   8400
      Picture         =   "frmInPatientMedicine.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5760
      Width           =   1215
   End
   Begin VB.TextBox txtBillID 
      Height          =   285
      Left            =   7920
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save"
      Height          =   855
      Left            =   3120
      Picture         =   "frmInPatientMedicine.frx":5C7D
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   10080
      Width           =   1215
   End
   Begin VB.TextBox txtgrndtot 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   9360
      Width           =   1095
   End
   Begin VB.TextBox txtdisgvn 
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   9360
      Width           =   1575
   End
   Begin VB.TextBox txtpayable 
      Height          =   375
      Left            =   8280
      TabIndex        =   2
      Top             =   9360
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CLOSE"
      Height          =   855
      Left            =   6000
      Picture         =   "frmInPatientMedicine.frx":612B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10080
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPIssue 
      Height          =   285
      Left            =   7920
      TabIndex        =   28
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      _Version        =   393216
      Format          =   47448065
      CurrentDate     =   38353
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
      FormHeightDT    =   11655
      FormWidthDT     =   10140
      FormScaleHeightDT=   11145
      FormScaleWidthDT=   10020
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   2895
      Left            =   120
      Top             =   2640
      Width           =   9735
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Medicine Type"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   38
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6360
      TabIndex        =   37
      Top             =   4920
      Width           =   1140
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount Given"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6360
      TabIndex        =   36
      Top             =   4440
      Width           =   1290
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6360
      TabIndex        =   35
      Top             =   3960
      Width           =   660
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Per Unit"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   34
      Top             =   4920
      Width           =   1125
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6360
      TabIndex        =   33
      Top             =   3480
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Date"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6360
      TabIndex        =   32
      Top             =   3000
      Width           =   930
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Medicine Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   31
      Top             =   3840
      Width           =   1290
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Medicine ID"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   30
      Top             =   3240
      Width           =   1005
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Units In Stock"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   29
      Top             =   4320
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name of Patient"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admission Code"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   15
      Top             =   1440
      Width           =   1380
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      TabIndex        =   14
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      TabIndex        =   13
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IN PATIENTS MEDICINE ISSUE "
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
      Left            =   2040
      TabIndex        =   12
      Top             =   360
      Width           =   6300
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   9480
      Width           =   1095
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Discount Given"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   10
      Top             =   9480
      Width           =   1335
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amount Payable"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      TabIndex        =   9
      Top             =   9480
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   1095
      Left            =   6120
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   3735
   End
End
Attribute VB_Name = "frmInPatientMedicine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Row As Integer
Dim stock As Integer




Private Sub cmbAdmitID_Click()
Dim rsSearch As Recordset
Set rsSearch = New ADODB.Recordset
Dim rsName As Recordset
Set rsName = New ADODB.Recordset

rsSearch.Open " Select * from Admission_Details where Admission_ID = '" & cmbAdmitID & " '", cnPatients, adOpenDynamic, adLockReadOnly
If rsSearch.RecordCount = 1 Then
    rsName.Open "Select * from In_Patient_Details where Patient_ID = '" & rsSearch(1) & "'", cnPatients, adOpenDynamic, adLockReadOnly
        If rsName.RecordCount = 1 Then
            txtName = rsName(1) & " " & rsName(2)
        Else
            MsgBox "Database Error", vbCritical
            rsName.Close
            Exit Sub
        End If
Else
    MsgBox "Database Error", vbCritical
    rsSearch.Close
    Exit Sub
End If
rsSearch.Close
rsName.Close


End Sub

Private Sub cmbMedID_Click()

Dim rsAddMedName As Recordset
Set rsAddMedName = New ADODB.Recordset


rsAddMedName.Open "Select * from Medicine_Details where ProductID = '" & cmbMedID & "'", cnPatients, adOpenDynamic, adLockReadOnly


If rsAddMedName.RecordCount > 1 Then
    MsgBox " Database Error"
    stock = 0
    txtStock = stock
    Exit Sub
ElseIf rsAddMedName.RecordCount = 0 Then
    txtmedname = ""
    txtRPU = "0.00"
    stock = 0

Else
    txtmedname = rsAddMedName(1)
    txtRPU = rsAddMedName(5)
    stock = rsAddMedName(4)
    
End If

txtStock = stock

rsAddMedName.Close




End Sub

Private Sub cmbMedType_Click()
Dim rsAddMedID As Recordset
Set rsAddMedID = New ADODB.Recordset
cmbMedID.clear

rsAddMedID.Open "Select * from Medicine_Details where CategoryID = '" & cmbMedType & "'", cnPatients, adOpenDynamic, adLockReadOnly

If rsAddMedID.EOF = False Then
    rsAddMedID.MoveFirst

    While rsAddMedID.EOF = False
        cmbMedID.AddItem rsAddMedID(0)
        cmbMedID.Text = rsAddMedID(0)
        rsAddMedID.MoveNext
    Wend
   

End If
If cmbMedID.ListCount = 0 Then
    txtmedname = ""
    txtRPU = "0"
    txtStock = "0"
End If

rsAddMedID.Close

End Sub



Private Sub cmdAddList_Click()
Dim rsChkPatient As Recordset
Set rsChkPatient = New ADODB.Recordset

rsChkPatient.Open "select * from In_Patient_Discharge where Admission_ID = '" & cmbAdmitID & "'", cnPatients, adOpenDynamic, adLockReadOnly
If rsChkPatient.EOF = False Then
    MsgBox "The Patient has already discharged", vbCritical
    Exit Sub
End If
rsChkPatient.Close


'On Error Resume Next
    Dim rsMed As Recordset
    Dim i As Integer
    

    
    
    
If MFG.Rows > 2 Then
    For i = 1 To MFG.Rows - 2 Step 1
        If MFG.TextMatrix(i, 2) = cmbMedID Then
            MsgBox "Medicine Already Exist In The List Cannot Add Same Medicine Again.....", vbCritical + vbOKOnly
            Exit Sub
        End If
    Next i
End If




If txtAmount = "" Or txttotamt = "" Or txtqty = "" Or txtRPU = "" Then
    MsgBox "Please Enter the relevant Fields", vbCritical
    Exit Sub
End If
If Val(txtqty) = 0 Then
    MsgBox "Quantity Cannot be Zero", vbCritical
    txtqty.SetFocus
    SendKeys "{Home}+{End}"
    Exit Sub
End If
If Val(txtqty) > stock Then
    MsgBox "The Quantity Cannot be greater than Stock", vbCritical
    txtqty.SetFocus
    SendKeys "{Home}+{End}"
    Exit Sub
End If

Row = MFG.Rows - 1
With MFG

        .Rows = .Rows + 1
                
        MFG.TextMatrix(Row, 1) = txtBillID
        MFG.TextMatrix(Row, 2) = cmbMedID
        MFG.TextMatrix(Row, 3) = txtmedname
        MFG.TextMatrix(Row, 4) = DTPIssue
        MFG.TextMatrix(Row, 5) = txtqty
        MFG.TextMatrix(Row, 6) = txtRPU
        MFG.TextMatrix(Row, 7) = txtdis
        MFG.TextMatrix(Row, 8) = txttotamt

  
    .FixedRows = 1
    .RowHeight(0) = .RowHeight(1) * 1.5
     Functions.SizeColumns MFG, Me
     MFGVALUES
     
     Row = Row + 1
     
End With



Call CalcFinal
Call TextClear

End Sub
Private Sub CalcFinal()


Dim amount As Double
Dim Discount As Double
Dim Total As Double

If MFG.Rows > 2 Then
    For i = 1 To MFG.Rows - 2 Step 1
        amount = amount + (Val(MFG.TextMatrix(i, 6)) * Val(MFG.TextMatrix(i, 5)))
        Discount = Discount + MFG.TextMatrix(i, 7)
        Total = Total + MFG.TextMatrix(i, 8)
        
    Next i
End If
txtgrndtot = amount
txtdisgvn = Discount
txtpayable = Total

txtStock = Val(txtStock) - Val(txtqty)
Debug.Print Val(amount) - Val(Discount)

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


Dim rsOrderID As Recordset
    Dim OID As String
    Set rsOrderID = New ADODB.Recordset
    ' Generatin Order Details ID
    OID = Functions.UID(6, "MODRID_")
    rsOrderID.Open " Select * from InPatient_Orders", cnPatients, adOpenDynamic, adLockPessimistic
    If rsOrderID.EOF = False Then
    While rsOrderID.EOF = False
        If rsOrderID(0) = PID Then
            OID = Functions.UID(6, "MODRID_")
            rsOrderID.MoveFirst
        End If
    rsOrderID.MoveNext
    Wend
    End If
    
Dim flag, flag1, flag2 As Boolean
flag = False
flag1 = False
flag2 = False
    
    
    rsOrderID.AddNew
        rsOrderID(0) = OID
        rsOrderID(1) = cmbAdmitID
        rsOrderID(2) = DTPDate
        
    rsOrderID.Update
    flag = True
    rsOrderID.Close
    
Dim rsMed As Recordset
Set rsMed = New ADODB.Recordset
Dim MID As String
Dim rsAddPatient As Recordset
Set rsAddPatient = New ADODB.Recordset
   
Dim rsStock As Recordset
Set rsStock = New ADODB.Recordset

rsMed.Open "SELECT * FROM InPatient_Order_Details", cnPatients, adOpenDynamic, adLockPessimistic
For i = 1 To MFG.Rows - 2 Step 1

    ' Generatin Order Details ID
    MID = Functions.UID(6, "ODRDTL_")
    rsAddPatient.Open " Select * from InPatient_Order_Details", cnPatients, adOpenDynamic, adLockReadOnly
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
                !OrderID = OID
                !ProductID = MFG.TextMatrix(i, 2)
                !DateSold = Format(MFG.TextMatrix(i, 4), "Short Date")
                !QUANTITY = Val(MFG.TextMatrix(i, 5))
                !UNITPRICE = Val(MFG.TextMatrix(i, 6))
                !Discount = Val(MFG.TextMatrix(i, 7))
            .Update
            flag1 = True
        End With
        rsStock.Open "select * from Medicine_Details where ProductID= '" & MFG.TextMatrix(i, 2) & "'", cnPatients, adOpenDynamic, adLockPessimistic
        If rsStock.EOF = False Then
            rsStock(4) = rsStock(4) - Val(MFG.TextMatrix(i, 5))
            rsStock.Update
            flag2 = True
        End If
        rsStock.Close

Next
rsMed.Close

If flag = True And flag1 = True And flag2 = True Then
    MsgBox "Record Saved Succesfully !!", vbInformation
Else
    MsgBox "An Error Occured while attempting to update the database", vbCritical, "Record Save Error"
End If

Unload Me


End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call Functions.DisableMenu
Me.WindowState = vbMaximized



Call AddInPatientDetails
Call AddMedicineDetails
Call GenerateBillID
Call MFGVALUES

DTPDate = Date
DTPIssue = Date
stock = 0
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 22 Then KeyAscii = 0: Exit Sub
    KeyAscii = DataEntryValidation(KeyAscii, ActiveControl.Tag)
End Sub
Public Sub AddInPatientDetails()
Dim rsAddPatient As Recordset
Set rsAddPatient = New ADODB.Recordset

rsAddPatient.Open "Select * from Admission_Details", cnPatients, adOpenDynamic, adLockReadOnly

If rsAddPatient.EOF = False Then
rsAddPatient.MoveFirst

While rsAddPatient.EOF = False
    cmbAdmitID.AddItem rsAddPatient(0)
    cmbAdmitID.Text = rsAddPatient(0)
    rsAddPatient.MoveNext
Wend


End If

rsAddPatient.Close



End Sub

Public Sub AddMedicineDetails()
Dim rsAddMed As Recordset
Set rsAddMed = New ADODB.Recordset

rsAddMed.Open "Select * from Medicine_Categories", cnPatients, adOpenDynamic, adLockReadOnly

If rsAddMed.EOF = False Then
rsAddMed.MoveFirst

While rsAddMed.EOF = False
    cmbMedType.AddItem rsAddMed(0)
    cmbMedType.Text = rsAddMed(0)
    rsAddMed.MoveNext
Wend

End If

rsAddMed.Close





End Sub

Public Sub GenerateBillID()

    Dim rsAddPatient As Recordset
    Dim MID As String
    Set rsAddPatient = New ADODB.Recordset
  
    MID = Functions.UID(6, "MedID_")
    rsAddPatient.Open " Select * from InPatient_Orders", cnPatients, adOpenKeyset, adLockReadOnly
    While rsAddPatient.EOF = False
        If rsAddPatient(0) = MID Then
            MID = Functions.UID(6, "MedID_")
            rsAddPatient.MoveFirst
        End If
    rsAddPatient.MoveNext
    Wend
    rsAddPatient.Close
    txtBillID = MID


End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Functions.EnableMenu
End Sub

Private Sub txtdis_Change()
txttotamt = Val(txtAmount) - Val(txtdis)
End Sub

Private Sub txtqty_Change()

txtAmount = Val(txtRPU) * Val(txtqty)

End Sub

Public Sub MFGVALUES()
MFG.TextMatrix(0, 1) = "ORDER ID"
MFG.TextMatrix(0, 2) = "MEDICINE CODE"
MFG.TextMatrix(0, 3) = "MEDICINE NAME"
MFG.TextMatrix(0, 4) = "DATE OF ISSUE"
MFG.TextMatrix(0, 5) = "QUANTITY"
MFG.TextMatrix(0, 6) = "UNIT PRICE"
MFG.TextMatrix(0, 7) = "DISCOUNT"
MFG.TextMatrix(0, 8) = "TOTAL AMOUNT"
Functions.SizeColumnHeaders MFG, Me

End Sub

Public Sub TextClear()
txtqty = ""
txtdis = ""
txtAmount = ""
txttotamt = ""
txtStock = ""
txtStock = stock - Val(txtqty)
cmbMedID_Click
End Sub

Private Sub txtqty_LostFocus()
If Val(txtqty) > stock Then
    MsgBox "The Quantity Cannot be greater than Stock", vbCritical
End If

End Sub
