VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frm_employee 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Management System"
   ClientHeight    =   4710
   ClientLeft      =   2775
   ClientTop       =   3585
   ClientWidth     =   10605
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_employee.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   10605
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   600
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      AutoCenterForm  =   -1  'True
      FormHeightDT    =   5190
      FormWidthDT     =   10695
      FormScaleHeightDT=   4710
      FormScaleWidthDT=   10605
   End
   Begin VB.CommandButton cmd_doctor 
      Appearance      =   0  'Flat
      Caption         =   "Doctor Section"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   2280
      Width           =   3495
   End
   Begin VB.CommandButton cmd_salary 
      Appearance      =   0  'Flat
      Caption         =   "Salary managment"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   2880
      Width           =   3495
   End
   Begin VB.CommandButton cmd_employee 
      Appearance      =   0  'Flat
      Caption         =   "Employee Section"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   1680
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Settings"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   3600
      Width           =   3495
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize2 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      FormHeightDT    =   5190
      FormWidthDT     =   10695
      FormScaleHeightDT=   4710
      FormScaleWidthDT=   10605
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Height          =   2895
      Left            =   2760
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   5535
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000009&
      BorderWidth     =   5
      Height          =   855
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   10095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Managment System"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   10575
   End
End
Attribute VB_Name = "frm_employee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3CBB3059014C"
Option Explicit

'##ModelId=3CBB30590155
Private str_choice As String
'##ModelId=3CBB30590158
Dim con As ADODB.Connection
'##ModelId=3CBB30590161
Dim rs As ADODB.Recordset
'##ModelId=3CBB30590164
Dim rs1 As ADODB.Recordset
'##ModelId=3CBB3059016B
Dim rs2 As ADODB.Recordset
'##ModelId=3CBB3059016C
Dim ch
Dim str_choice_1

'##ModelId=3CBB30590173
Private Sub cmd_add_Click()
str_choice_1 = "add"

If str_choice = "emp" Then
frm_add_employee.Show
ElseIf str_choice = "doc" Then
'frm_add_doctor.Show
Else
frm_add_salary_info.Show
End If
Unload Me
End Sub

'##ModelId=3CBB30590174
Private Sub cmd_delete_Click()

End Sub


'##ModelId=3CBB3059017D
Private Sub cmd_add_doc_Click()
str_choice_1 = "add"
'frm_add_doctor.Show
End Sub

'##ModelId=3CBB3059017E
Private Sub cmd_add_emp_Click()
str_choice_1 = "add"
frm_add_employee.Show
End Sub

'##ModelId=3CBB30590187
Private Sub cmd_add_sal_Click()
str_choice_1 = "add"
frm_add_salary_info.Show
End Sub

Private Sub cmd_back_to_main_Click()
Unload Me
End Sub

''##ModelId=3CBB305901B9
Private Sub cmd_doctor_Click()
frm_app_count.Show
End Sub

'##ModelId=3CBB305901BA
Private Sub cmd_employee_Click()
frm_add_employee.Show
End Sub

'##ModelId=3CBB305901C3
Private Sub cmd_salary_Click()
frm_add_salary_info.Show
End Sub

'##ModelId=3CBB305901C4
Private Sub Command1_Click()
frm_settngs.Show
End Sub

'##ModelId=3CBB305901CD
Private Sub Form_Load()
'Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
'con.CursorLocation = adUseClient
'con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\pay\Pay.mdb;persist security info=false"



End Sub


