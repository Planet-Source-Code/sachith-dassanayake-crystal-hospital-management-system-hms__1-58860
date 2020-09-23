VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frm_settngs 
   BackColor       =   &H00FF8080&
   Caption         =   "Settings"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_settngs.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   7680
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "CATEOGRY DETAILS"
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
      Height          =   3615
      Left            =   720
      TabIndex        =   16
      Top             =   4200
      Width           =   6255
      Begin VB.TextBox txt_cat 
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Top             =   1080
         Width           =   3135
      End
      Begin VB.CommandButton cmd_cat_add 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   960
         Picture         =   "frm_settngs.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton cmd_cat_del 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2640
         Picture         =   "frm_settngs.frx":5C69
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton cmd_cat_viewall 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&View All"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4320
         Picture         =   "frm_settngs.frx":6122
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox txt_cat_id 
         Height          =   375
         Left            =   2640
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.CommandButton cmd_end1 
         Caption         =   "...."
         Height          =   375
         Left            =   3960
         TabIndex        =   7
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Category Name"
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
         Left            =   600
         TabIndex        =   18
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label6 
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
         Height          =   495
         Left            =   600
         TabIndex        =   17
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "DEPARTMENT DETAILS"
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
      Height          =   3615
      Left            =   720
      TabIndex        =   12
      Top             =   360
      Width           =   6255
      Begin VB.TextBox txt_department 
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   1320
         Width           =   3135
      End
      Begin VB.CommandButton cmd_dep_add 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   960
         Picture         =   "frm_settngs.frx":65DD
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmd_dept_delete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2640
         Picture         =   "frm_settngs.frx":6A64
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmd_dept_view 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&View All"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4320
         Picture         =   "frm_settngs.frx":6F1D
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txt_dep_id 
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         TabIndex        =   0
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmd_end 
         Caption         =   "...."
         Height          =   375
         Left            =   3960
         TabIndex        =   1
         Top             =   720
         Width           =   375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Department Name"
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
         Left            =   600
         TabIndex        =   15
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Department ID"
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
         Left            =   600
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      Picture         =   "frm_settngs.frx":73D8
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8040
      Width           =   1215
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
      FormHeightDT    =   9690
      FormWidthDT     =   7800
      FormScaleHeightDT=   9180
      FormScaleWidthDT=   7680
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
End
Attribute VB_Name = "frm_settngs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3CBB305B01B2"
'##ModelId=3CBB305B01BC
Private intdep As Integer
'##ModelId=3CBB305B01C6
Private intcat As Integer
'##ModelId=3CBB305B01C7
Private strok As String
'##ModelId=3CBB305B01D0
Private strno As String

'##ModelId=3CBB305B01D1
Private Sub cmd_cat_add_Click()
Dim a
strok = "ok"

If txt_cat.Text = "" Then
a = MsgBox("Please Enter Data To Procede", vbCritical, "Error")
Else
rscat.AddNew
rscat.Fields(0) = txt_cat_id.Text
rscat.Fields(1) = txt_cat.Text
rscat.Fields(2) = strok
rscat.Update

Call cat_no
txt_cat_id = intcat
txt_cat = ""
End If
End Sub

Private Sub cmd_cat_del_Click()
Dim res As String
If txt_cat_id = "" Then
MsgBox "There Is No Current Record", vbInformation
Else
res = MsgBox("Do You Want To Delete The Current Record ? ", vbCritical + vbYesNo, "Data Deletion")
If res = vbYes Then
rsdep.Close
rsdep.Open "Select * From Doctor_special where id='" & txt_cat_id & "'", con, adOpenDynamic, adLockOptimistic
txt_cat = rsdep.Fields(1)
con.Execute ("delete from Doctor_special where id='" & txt_cat_id & "'")
con.Execute ("insert into Doctor_special values('" & txt_cat_id & "','" & txt_cat & "','" & strno & "')")
txt_cat = ""
intcat = 0
Call cat_no
txt_cat_id.Text = intdep
ElseIf res = vbNo Then
MsgBox "Deletion Cancled", vbInformation
End If
End If
End Sub

'##ModelId=3CBB305B01DA
Private Sub cmd_cat_viewall_Click()
frm_view_docat.Show
End Sub

'##ModelId=3CBB305B01DB
Private Sub cmd_dep_add_Click()
Dim a
If txt_department.Text = "" Then
a = MsgBox("Please Enter Data To Procede", vbCritical, "Error")
Else
rsdep.AddNew
rsdep.Fields(0) = txt_dep_id.Text
rsdep.Fields(1) = txt_department.Text
rsdep.Fields(2) = strok
rsdep.Update
intdep = 0
Call dep_no
txt_dep_id = intdep
txt_department = ""
End If

End Sub

'##ModelId=3CBB305B01E4
Private Sub cmd_dept_delete_Click()

Dim res As String
If txt_dep_id = "" Then
MsgBox "There Is No Current Record", vbInformation
Else
res = MsgBox("Do You Want To Delete The Current Record ? ", vbCritical + vbYesNo, "Data Deletion")
If res = vbYes Then
rsdep.Close
rsdep.Open "Select * From Department where dept_id='" & txt_dep_id & "'", con, adOpenDynamic, adLockOptimistic
txt_department = rsdep.Fields(1)
con.Execute ("delete from department where dept_id='" & txt_dep_id & "'")
con.Execute ("insert into department values('" & txt_dep_id & "','" & txt_department & "','" & strno & "')")
txt_department = ""
intdep = 0
Call dep_no
txt_dep_id.Text = intdep

ElseIf res = vbNo Then
MsgBox "Deletion Cancled", vbInformation
End If
End If

End Sub

'##ModelId=3CBB305B01E5
Private Sub cmd_dept_view_Click()
frm_view_dep.Show
Call dep_no
End Sub

'##ModelId=3CBB305B01EE
Private Sub cmd_end_Click()
txt_dep_id.Enabled = True
txt_dep_id.Text = "  "
txt_dep_id.SetFocus

End Sub


'##ModelId=3CBB305B01EF
Private Sub cmd_end1_Click()
txt_cat_id.Enabled = True
txt_cat_id.Text = " "
txt_cat_id.SetFocus

End Sub

'##ModelId=3CBB305B01F8
Private Sub Command1_Click()
Unload Me
frm_employee.Show
End Sub

'##ModelId=3CBB305B01F9
Private Sub Form_Load()
Dim rsDocs As New Recordset
strok = "ok"
strno = "no"
Set con = New ADODB.Connection
Set rsdep = New ADODB.Recordset
Set rscat = New ADODB.Recordset

rsdep.Open "select * from Department", cnPatients, adOpenDynamic, adLockOptimistic
Call dep_no
Call cat_no
txt_dep_id = intdep
txt_cat_id = intcat
End Sub

Private Sub txt_cat_Change()
If txt_cat_id = "" Then
MsgBox "Plaese Enter a ID To Continue", vbInformation, "Alert"
ElseIf Not txt_cat_id = "" Then
rsdep.Close
rsdep.Open "select * from Doctor_special where id='" & txt_cat_id & "' "
If rsdep.RecordCount = 0 Then
cmd_cat_add.Enabled = True
cmd_cat_del.Enabled = False
Else
cmd_cat_add.Enabled = False
cmd_cat_del.Enabled = True
End If
End If
End Sub


Public Sub dep_no()
Dim fnum As Integer
'rsdep.Close
'rsdep.Open "select * from Department  ", con, adOpenDynamic, adLockOptimistic
'rsdep.MoveFirst
'fnum = 101
'intdep = rsdep.RecordCount
'intdep = intdep + 2
'intdep = intdep + fnum
End Sub

Public Sub cat_no()
Dim dnum As Integer
'rscat.Close
'rscat.Open "select * from Doctor_special   ", cnPatients, adOpenDynamic, adLockOptimistic
'rsdep.MoveFirst
'dnum = 101
'MsgBox (intcat)
'intcat = rsdep.RecordCount
'MsgBox (intcat)
'intcat = intcat + 2
'MsgBox (intcat)
'intcat = intcat + dnum
'MsgBox (intcat)
End Sub
