VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frm_add_employee 
   BackColor       =   &H00FF8080&
   Caption         =   "Employee Data Entry"
   ClientHeight    =   10770
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10770
   ScaleWidth      =   9150
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Employee Details"
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
      Height          =   6735
      Left            =   960
      TabIndex        =   9
      Top             =   960
      Width           =   7335
      Begin VB.ComboBox cmb_emp_department 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3480
         TabIndex        =   18
         Top             =   4320
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "Female"
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
         Index           =   1
         Left            =   4560
         TabIndex        =   17
         Top             =   3840
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF8080&
         Caption         =   "Male"
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
         Index           =   0
         Left            =   3480
         TabIndex        =   16
         Top             =   3840
         Width           =   855
      End
      Begin VB.TextBox txt_emp_bsal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   15
         Top             =   5880
         Width           =   1215
      End
      Begin VB.TextBox txt_emp_insurno 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   14
         Top             =   5400
         Width           =   1215
      End
      Begin VB.TextBox txt_emp_insurecorp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   13
         Top             =   4920
         Width           =   2775
      End
      Begin VB.TextBox txt_emp_address 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   3480
         TabIndex        =   12
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txt_emp_name 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   11
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txt_empid 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3480
         TabIndex        =   10
         Top             =   480
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   3480
         TabIndex        =   19
         Top             =   3240
         Width           =   1455
         _ExtentX        =   2566
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
         Format          =   45744129
         CurrentDate     =   38338
      End
      Begin MSMask.MaskEdBox txt_emp_telephone 
         Height          =   285
         Left            =   3480
         TabIndex        =   20
         ToolTipText     =   "Enter the account number here"
         Top             =   2760
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
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
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Basic Salary"
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
         Left            =   1200
         TabIndex        =   30
         Top             =   5880
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance Number"
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
         Left            =   1200
         TabIndex        =   29
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance Coperation"
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
         Left            =   1200
         TabIndex        =   28
         Top             =   4920
         Width           =   2175
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Sex"
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
         Left            =   1200
         TabIndex        =   27
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
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
         Left            =   1200
         TabIndex        =   26
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Birth"
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
         Left            =   1200
         TabIndex        =   25
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone"
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
         Left            =   1200
         TabIndex        =   24
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Left            =   1200
         TabIndex        =   23
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
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
         Left            =   1200
         TabIndex        =   22
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID"
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
         Left            =   1200
         TabIndex        =   21
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Operations"
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
      Height          =   2655
      Left            =   720
      TabIndex        =   0
      Top             =   7920
      Width           =   7695
      Begin VB.CommandButton cmd_emp_add 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         Height          =   855
         Left            =   360
         Picture         =   "Main.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmd_emp_save 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   855
         Left            =   1800
         Picture         =   "Main.frx":12F5
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmd_emp_modify 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Modify"
         Height          =   855
         Left            =   3240
         Picture         =   "Main.frx":17A3
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmd_emp_viewall 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&View All"
         Height          =   855
         Left            =   3240
         Picture         =   "Main.frx":1C7A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmd_emp_delete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   855
         Left            =   6120
         Picture         =   "Main.frx":2135
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmd_emp_search 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search"
         CausesValidation=   0   'False
         Height          =   855
         Left            =   4680
         Picture         =   "Main.frx":25F8
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmd_employee_refresh 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Refresh"
         Height          =   855
         Left            =   1800
         Picture         =   "Main.frx":2AD0
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CommandButton cmd_emp_back 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   855
         Left            =   4680
         Picture         =   "Main.frx":2F76
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1440
         Width           =   1215
      End
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
      FormHeightDT    =   11280
      FormWidthDT     =   9270
      FormScaleHeightDT=   10770
      FormScaleWidthDT=   9150
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE DETAILS"
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
      Left            =   2880
      TabIndex        =   31
      Top             =   120
      Width           =   3990
   End
End
Attribute VB_Name = "frm_add_employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3CBB30590295"
Private rsDocs As New ADODB.Recordset

'##ModelId=3CBB3059029F
Private ddate As Date
'##ModelId=3CBB305902A0
Private Sub cmd_emp_add_Click()
Dim ctl As Control

For Each ctl In Controls
    If TypeOf ctl Is TextBox Then
        ctl.Locked = False
    End If
    If TypeOf ctl Is ComboBox Then
        ctl.Locked = False
    End If
Next



cmd_emp_add.Enabled = False
cmd_emp_save.Enabled = True

Call clear
End Sub

'##ModelId=3CBB305902A9
Private Sub cmd_emp_back_Click()
Unload Me
frm_employee.Show
End Sub

'##ModelId=3CBB305902AA
Private Sub cmd_emp_delete_Click()
If txt_empid = "" Then
MsgBox "There Is No Current Record", vbInformation
Else
res = MsgBox("Do You Want To Delete The Current Record ? ", vbCritical + vbYesNo, "Data Deletion")
If res = vbYes Then
cnPatients.Execute ("delete from employee where emp_id='" & txt_empid & "'")
Call clear
ElseIf res = vbNo Then
MsgBox "Deletion Cancled", vbInformation
End If
End If
End Sub

'##ModelId=3CBB305902B3
Private Sub cmd_emp_modify_Click()
txt_empid.Enabled = False

If txt_empid = "" Then
MsgBox "There Is No Current Record", vbInformation
Else
res = MsgBox("Do You Want To Modify The Current Record ? ", vbCritical + vbYesNo, "Data Modification")
If res = vbYes Then

cnPatients.Execute ("delete from employee where emp_id='" & txt_empid & "'")
cnPatients.Execute ("Insert into employee values('" & txt_empid & "','" & txt_emp_name & "','" & txt_emp_address & "','" & txt_emp_telephone & "','" & DTPicker1.Value & "','" & sex & "','" & cmb_emp_department.Text & "','" & txt_emp_insurecorp & "','" & txt_emp_insurno & "','" & txt_emp_bsal & "')")

Call clear
ElseIf res = vbNo Then
MsgBox "Modifcation Cancled", vbInformation
End If
End If

End Sub

'##ModelId=3CBB305902B4
Private Sub cmd_emp_save_Click()
cmd_emp_save.Enabled = False
cmd_emp_add.Enabled = True
Dim sex As String
Dim x As String
x = "0"

If Option1(0).Value = True Then
sex = "male"
Else
sex = "female"
End If

cnPatients.Execute ("Insert into employee values('" & txt_empid & "','" & txt_emp_name & "','" & txt_emp_address & "','" & txt_emp_telephone & "','" & ddate & "','" & sex & "','" & cmb_emp_department.Text & "','" & txt_emp_insurecorp & "','" & txt_emp_insurno & "','" & txt_emp_bsal & "')")
cnPatients.Execute ("Insert into slip values('" & txt_empid & "','" & txt_emp_name & "','" & cmb_emp_department.Text & "','" & txt_emp_address & "','" & txt_emp_bsal & "','" & x & "','" & x & "','" & x & "','" & x & "','" & x & "','" & x & "','" & x & "','" & x & "','" & x & "','" & x & "','" & x & "')")

Call clear
End Sub

Private Sub cmd_emp_search_Click()
txt_empid.Enabled = False
str_search_number = InputBox("Enter The Employee Number", "Data Search")

rsDocs.MoveFirst
While Not rsDocs.EOF
If rsDocs.Fields(0) = str_search_number Then
MsgBox "Record Found"

Call recassign
End If
rsDocs.MoveNext
Wend


txt_empid.Enabled = False
End Sub

Private Sub cmd_emp_viewall_Click()
frm_employee_view.Show
End Sub

Private Sub cmd_employee_refresh_Click()
rsDocs.Close

rsDocs.Open "select * from employee where emp_id='" & txt_empid & "' ", cnPatients, adOpenDynamic, adLockOptimistic

If rsDocs.RecordCount = 0 Then
MsgBox "No Current Records To Refresh", vbInformation, "Alert"

ElseIf rsDocs.RecordCount > 0 Then
With rsDocs
.Requery
.MoveFirst
End With
End If

End Sub

'##ModelId=3CBB305902C8
Private Sub Form_Load()

Dim ctl As Control

For Each ctl In Controls
    If TypeOf ctl Is TextBox Then
        ctl.Locked = True
    End If
    If TypeOf ctl Is ComboBox Then
        ctl.Locked = True
    End If
Next
    




Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset

Set rsadd = New ADODB.Recordset
Set rsmod = New ADODB.Recordset
Set rsdel = New ADODB.Recordset

con.CursorLocation = adUseClient
'con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\pay\Pay.mdb;persist security info=false"
strst = "del"
Dim dep

Set rsDocs = New ADODB.Recordset

rsDocs.Open "select * from employee  ", cnPatients, adOpenDynamic, adLockOptimistic
rs1.Open "select * from Department", cnPatients, adOpenDynamic, adLockOptimistic
rs.Open "select * from slip", cnPatients, adOpenDynamic, adLockOptimistic

rs1.MoveFirst
While rs1.EOF = False
cmb_emp_department.AddItem rs1!dept_na
rs1.MoveNext
Wend

If rsDocs.RecordCount > 0 Then
rsDocs.MoveFirst
End If


cmd_emp_save.Enabled = False
ddate = DTPicker1.Value
End Sub

'##ModelId=3CBB305902D1
Public Sub recassign()
txt_empid.Text = rsDocs.Fields(0)
txt_emp_name.Text = rsDocs.Fields(1)
txt_emp_address.Text = rsDocs.Fields(2)
txt_emp_telephone.Text = rsDocs.Fields(3)
DTPicker1.Value = rsDocs.Fields(4)
cmb_emp_department.Text = rsDocs.Fields(6)
txt_emp_insurecorp.Text = rsDocs.Fields(7)
txt_emp_insurno.Text = rsDocs.Fields(8)
txt_emp_bsal.Text = rsDocs.Fields(9)

If rsDocs.Fields(5) = "male" Then
Option1(0).Value = True
Option1(1).Value = False
ElseIf rsDocs.Fields(5) = "female" Then
Option1(0).Value = False
Option1(1).Value = True
End If

End Sub

'##ModelId=3CBB305902D2
Public Sub clear()
With frm_add_employee
.txt_emp_address = ""
.txt_emp_bsal = ""
.DTPicker1.Value = ddate
.txt_emp_insurecorp = ""
.txt_emp_insurno = ""
.txt_emp_name = ""
.txt_emp_telephone = ""
.txt_empid = ""
.cmb_emp_department = ""
End With
End Sub

