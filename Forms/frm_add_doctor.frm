VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm_add_doctor 
   BackColor       =   &H80000001&
   Caption         =   "Doctor Data Entry"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7710
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmd_doc_search 
      Appearance      =   0  'Flat
      Caption         =   "Search"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   30
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmd_doc_refresh 
      Appearance      =   0  'Flat
      Caption         =   "Refresh"
      Height          =   375
      Left            =   2160
      TabIndex        =   29
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmd_doc_save 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   375
      Left            =   1800
      TabIndex        =   25
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmd_doc_viewall 
      Appearance      =   0  'Flat
      Caption         =   "&View All"
      Height          =   375
      Left            =   5880
      TabIndex        =   24
      Top             =   6480
      Width           =   1095
   End
   Begin VB.ComboBox cmb_status 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frm_add_doctor.frx":0000
      Left            =   2280
      List            =   "frm_add_doctor.frx":000A
      TabIndex        =   23
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton cmd_doc_modify 
      Appearance      =   0  'Flat
      Caption         =   "&Modify"
      Height          =   375
      Left            =   3120
      TabIndex        =   22
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmd_doc_delete 
      Appearance      =   0  'Flat
      Caption         =   "&Delete"
      Height          =   375
      Left            =   4080
      TabIndex        =   21
      Top             =   6960
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000001&
      Caption         =   "Female"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
   End
   Begin VB.CommandButton cmd_doc_back 
      Appearance      =   0  'Flat
      Caption         =   "<< &Back <<"
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   7440
      Width           =   6735
   End
   Begin VB.TextBox txt_docid 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox txt_doc_name 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   840
      Width           =   3735
   End
   Begin VB.TextBox txt_doc_address 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   1440
      Width           =   3735
   End
   Begin VB.TextBox txt_doc_insurecorp 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   4320
      Width           =   2775
   End
   Begin VB.TextBox txt_doc_insurno 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   4920
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H80000001&
      Caption         =   "Male"
      Height          =   255
      Index           =   0
      Left            =   2280
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.ComboBox cmb_doc_spec 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmd_add_doctor 
      Appearance      =   0  'Flat
      Caption         =   "&Add"
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   6480
      Width           =   975
   End
   Begin MSMask.MaskEdBox txt_Telephone 
      Height          =   405
      Left            =   2280
      TabIndex        =   26
      ToolTipText     =   "Enter the account number here"
      Top             =   2040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "###-#######"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox txt_mobile 
      Height          =   405
      Left            =   5520
      TabIndex        =   27
      ToolTipText     =   "Enter the account number here"
      Top             =   2040
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      HideSelection   =   0   'False
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "###-#######"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2280
      TabIndex        =   28
      Top             =   2640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   19529729
      CurrentDate     =   38338
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000004&
      BorderWidth     =   3
      Height          =   1815
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Width           =   7335
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Number"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4200
      TabIndex        =   20
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor ID"
      Height          =   495
      Left            =   360
      TabIndex        =   18
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Name"
      Height          =   495
      Left            =   360
      TabIndex        =   17
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone"
      Height          =   495
      Left            =   360
      TabIndex        =   15
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Of Birth"
      Height          =   495
      Left            =   360
      TabIndex        =   14
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Speciality"
      Height          =   495
      Left            =   360
      TabIndex        =   13
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Insurance Coperation"
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Insurance Number"
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   5640
      Width           =   1695
   End
End
Attribute VB_Name = "frm_add_doctor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3CBB305903D6"
Private strsex As String
Private ddate As Date


Private Sub cmd_doc_delete_Click()

If txt_docid = "" Then
MsgBox "There Is No Current Record", vbInformation
Else
res = MsgBox("Do You Want To Delete The Current Record ? ", vbCritical + vbYesNo, "Data Deletion")
If res = vbYes Then
con.Execute ("delete from doctor_detail where doc_id='" & txt_docid & "'")
Call clear
ElseIf res = vbNo Then
MsgBox "Deletion Cancled", vbInformation
End If
End If

End Sub

Private Sub cmd_doc_modify_Click()
txt_docid.Enabled = False

If txt_docid = "" Then
MsgBox "There Is No Current Record", vbInformation
Else
res = MsgBox("Do You Want To Modify The Current Record ? ", vbCritical + vbYesNo, "Data Modification")
If res = vbYes Then

con.Execute ("delete from doctor_detail where doc_id='" & txt_docid & "'")
con.Execute ("Insert into doctor_detail values('" & txt_docid & "','" & txt_doc_name & "','" & txt_doc_address & "','" & txt_Telephone & "','" & txt_mobile & "','" & DTPicker1.Value & "','" & strsex & "','" & cmb_doc_spec & "','" & txt_doc_insurecorp & "','" & txt_doc_insurno & "','" & cmb_status & "')")

Call clear
ElseIf res = vbNo Then
MsgBox "Modifcation Cancled", vbInformation
End If
End If
End Sub

Private Sub cmd_doc_refresh_Click()
rsDocs.Close
rsDocs.Open "select * from doctor_detail where doc_id='" & txt_docid & "' ", con, adOpenDynamic, adLockOptimistic

If rsDocs.RecordCount = 0 Then
MsgBox "No Current Records To Refresh", vbInformation, "Alert"

ElseIf rsDocs.RecordCount > 0 Then
With rsDocs
.Requery
.MoveFirst
End With
End If
End Sub

'##ModelId=3CBB305A0003
Private Sub cmd_doc_save_Click()
cmd_doc_save.Enabled = False
cmd_add_doctor.Enabled = True

Dim sex As String

If Option1(0).Value = True Then
strsex = "male"
Else
strsex = "female"
End If

con.Execute ("Insert into doctor_detail values('" & txt_docid & "','" & txt_doc_name & "','" & txt_doc_address & "','" & txt_Telephone & "','" & txt_mobile & "','" & DTPicker1.Value & "','" & strsex & "','" & cmb_doc_spec & "','" & txt_doc_insurecorp & "','" & txt_doc_insurno & "','" & cmb_status & "')")

Call clear
End Sub

Private Sub cmd_doc_search_Click()
rsDocs.Close
rsDocs.Open "select * from doctor_detail  ", con, adOpenDynamic, adLockOptimistic

txt_docid.Enabled = False

str_search_number = InputBox("Enter The Doctor ID ", "Data Search")
If rsDocs.RecordCount = 0 Then
MsgBox "No Current Records To Display", vbInformation, "Alert"
ElseIf rsDocs.RecordCount > 0 Then

rsDocs.MoveFirst
While Not rsDocs.EOF
If rsDocs.Fields(0) = str_search_number Then
MsgBox "Record Found"
Call recassign
End If
rsDocs.MoveNext
Wend
txt_docid.Enabled = False
End If
End Sub

'##ModelId=3CBB305A000C
Private Sub cmd_doc_viewall_Click()
frm_doctor_view.Show
End Sub

'##ModelId=3CBB305A000D
Private Sub Form_Load()

Set con = New ADODB.Connection
Set docs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset

Dim dep

cmd_add_doctor.Enabled = True
cmd_doc_save.Enabled = False
con.CursorLocation = adUseClient
con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\pay\Pay.mdb;persist security info=false"

rsDocs.Open "select * from doctor_detail  ", con, adOpenDynamic, adLockOptimistic
rs1.Open "select * from Doctor_special", con, adOpenDynamic, adLockOptimistic

rs1.MoveFirst
While rs1.EOF = False
cmb_doc_spec.AddItem rs1(1)
rs1.MoveNext
Wend
ddate = DTPicker1.Value
Call clear
End Sub
'##ModelId=3CBB305903E0
Private Sub cmd_add_doctor_Click()
cmd_doc_save.Enabled = True
cmd_add_doctor.Enabled = False
txt_doc_name.Enabled = True
txt_docid.Enabled = True
Call clear
End Sub

'##ModelId=3CBB305A0002
Private Sub cmd_doc_back_Click()
Unload Me
frm_employee.Show
End Sub
Public Sub clear()
 With frm_add_doctor
    .txt_doc_address.Text = ""
    .txt_doc_insurecorp.Text = ""
    .txt_doc_insurno.Text = ""
    .txt_doc_name.Text = ""
    .txt_docid.Text = ""
    .txt_mobile.Text = ""
    .txt_Telephone.Text = ""
    .cmb_doc_spec.Text = ""
    .cmb_status.Text = ""
    .DTPicker1.Value = ddate
   End With
End Sub

Public Sub recassign()
With frm_add_doctor
.txt_docid.Text = rsDocs.Fields(0)
.txt_doc_name.Text = rsDocs.Fields(1)
.txt_doc_address.Text = rsDocs.Fields(2)
.txt_Telephone.Text = rsDocs.Fields(3)
.txt_mobile.Text = rsDocs.Fields(4)
.DTPicker1.Value = rsDocs.Fields(5)
.cmb_doc_spec.Text = rsDocs.Fields(7)
.txt_doc_insurecorp.Text = rsDocs(8)
.txt_doc_insurno.Text = rsDocs.Fields(9)
.cmb_status.Text = rsDocs.Fields(10)

If rsDocs.Fields(6) = "male" Then
Option1(0).Value = True
Option1(1).Value = False
ElseIf rsDocs.Fields(6) = "female" Then
Option1(0).Value = False
Option1(1).Value = True
End If
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
rsDocs.Open
rsDocs.Close
End Sub
