VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frm_add_salary_info 
   BackColor       =   &H00FF8080&
   Caption         =   "Salary Mangment"
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   10095
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_add_salary_info.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   10095
   Begin VB.Frame Frame3 
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
      Height          =   7215
      Left            =   7800
      TabIndex        =   35
      Top             =   1200
      Width           =   2175
      Begin VB.CommandButton cmd_salary_add 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
         Height          =   735
         Left            =   480
         Picture         =   "frm_add_salary_info.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmd_salary_back 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   735
         Left            =   480
         Picture         =   "frm_add_salary_info.frx":5C91
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   6240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_sal_modify 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Modify"
         Height          =   735
         Left            =   480
         Picture         =   "frm_add_salary_info.frx":6195
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton cmd_sal_delete 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   735
         Left            =   480
         Picture         =   "frm_add_salary_info.frx":663A
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmd_sal_viewall 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&View All"
         Height          =   735
         Left            =   480
         Picture         =   "frm_add_salary_info.frx":6AF3
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmd_sal_save 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   735
         Left            =   480
         Picture         =   "frm_add_salary_info.frx":6FAE
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton cmd_sal_refresh 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Refresh"
         Height          =   735
         Left            =   480
         Picture         =   "frm_add_salary_info.frx":745C
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   5400
         Width           =   1095
      End
      Begin VB.CommandButton cmd_sal_search 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search"
         CausesValidation=   0   'False
         Height          =   735
         Left            =   480
         Picture         =   "frm_add_salary_info.frx":7902
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   2880
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   2535
      Left            =   360
      TabIndex        =   25
      Top             =   1200
      Width           =   7215
      Begin VB.TextBox Code 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2520
         TabIndex        =   30
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox add 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000080&
         Height          =   885
         Left            =   2520
         TabIndex        =   29
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox desig 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2520
         TabIndex        =   28
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox nm 
         Appearance      =   0  'Flat
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   2520
         TabIndex        =   27
         Top             =   720
         Width           =   3495
      End
      Begin VB.CommandButton cmd_brows 
         Appearance      =   0  'Flat
         Caption         =   "..."
         Height          =   255
         Left            =   3960
         TabIndex        =   26
         Top             =   360
         Width           =   375
      End
      Begin VB.Label label4 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   34
         Top             =   1080
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee's Code"
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
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   1665
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         Left            =   240
         TabIndex        =   32
         Top             =   720
         Width           =   555
      End
      Begin VB.Label label4 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   31
         Top             =   2160
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Salary Details"
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
      Height          =   4455
      Left            =   360
      TabIndex        =   0
      Top             =   3960
      Width           =   7215
      Begin VB.TextBox pt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         TabIndex        =   12
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox bp 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox da 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox hr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   9
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox ca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   8
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox ta 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox pf 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         MaxLength       =   5
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox ins 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         MaxLength       =   5
         TabIndex        =   5
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox it 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5520
         MaxLength       =   5
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox gpa 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         TabIndex        =   3
         Top             =   3600
         Width           =   1215
      End
      Begin VB.TextBox deduct 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         TabIndex        =   2
         Top             =   2640
         Width           =   1215
      End
      Begin VB.TextBox Net 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   1
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Basic Pay"
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
         Left            =   360
         TabIndex        =   24
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "D.A."
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
         Left            =   360
         TabIndex        =   23
         Top             =   840
         Width           =   390
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "H.R.A."
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
         Left            =   360
         TabIndex        =   22
         Top             =   1320
         Width           =   570
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C.C.A."
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
         Left            =   360
         TabIndex        =   21
         Top             =   1800
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transport Allowance"
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
         Left            =   360
         TabIndex        =   20
         Top             =   2280
         Width           =   2010
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "G.P.F"
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
         Left            =   4320
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance"
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
         Left            =   4320
         TabIndex        =   18
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Income Tax"
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
         Left            =   4320
         TabIndex        =   17
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P. Tax"
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
         Left            =   4320
         TabIndex        =   16
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gross Pay"
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
         Left            =   360
         TabIndex        =   15
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deductions"
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
         Left            =   360
         TabIndex        =   14
         Top             =   2760
         Width           =   1080
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nett Pay"
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
         Left            =   360
         TabIndex        =   13
         Top             =   3240
         Width           =   825
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
      FormHeightDT    =   9435
      FormWidthDT     =   10215
      FormScaleHeightDT=   8925
      FormScaleWidthDT=   10095
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE SALARY DETAILS"
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
      Left            =   2280
      TabIndex        =   44
      Top             =   360
      Width           =   5715
   End
End
Attribute VB_Name = "frm_add_salary_info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3CBB305A0156"
Option Explicit
'##ModelId=3CBB305A0162
Dim con As ADODB.Connection
'##ModelId=3CBB305A016C
Dim rs As ADODB.Recordset
'##ModelId=3CBB305A016D
Dim ch
'##ModelId=3CBB305A0174
Private Sub add_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  desig.SetFocus
End If
End Sub
'##ModelId=3CBB305A0177
Private Sub add_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 91 Or KeyAscii >= 97 And KeyAscii <= 123 Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 13 Then
    If Len(nm.Text) = 0 Or KeyAscii = 32 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 32 Then
          ch = "True"
        End If
    Else
        If ch = "true" Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            ch = "false"
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
    Else
        MsgBox ("Invalid Input")
        KeyAscii = 0
    End If
End Sub

'##ModelId=3CBB305A017F
Private Sub bp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  ca.SetFocus
End If
End Sub
'##ModelId=3CBB305A0188
Private Sub bp_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub
'##ModelId=3CBB305A018A
Private Sub bp_LostFocus()

If bp.Text <> "" Then
    da.Text = Val(bp.Text) * 0.458
    hr.Text = Val(bp.Text) * 0.3
    gpa.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text)
    deduct.Text = Val(pf.Text) + Val(ins.Text) + Val(it.Text) + Val(pt.Text)
    Net.Text = Val(gpa.Text) - Val(deduct.Text)
    End If
    If Code = "" Then
    cmd_salary_add.Enabled = False
    ElseIf nm = "" Then
    cmd_salary_add.Enabled = False
    ElseIf desig = "" Then
    cmd_salary_add.Enabled = False
    ElseIf add = "" Then
    cmd_salary_add.Enabled = False
    ElseIf bp = "" Then
    cmd_salary_add.Enabled = False
    Else
    cmd_salary_add.Enabled = True
    End If
End Sub
'##ModelId=3CBB305A0192
Private Sub ca_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    ta.SetFocus
End If
End Sub

'##ModelId=3CBB305A0195
Private Sub ca_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub
'##ModelId=3CBB305A019C
Private Sub ca_LostFocus()
    gpa.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text)
    Net.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) - Val(pf.Text) - Val(ins.Text) - Val(it.Text) - Val(pt.Text)
End Sub

Private Sub cmd_brows_Click()
frm_emp_list.Show
End Sub

Private Sub cmd_sal_delete_Click()
Dim res As String
If Code = "" Then
MsgBox "There Is No Current Record", vbInformation
Else
res = MsgBox("Do You Want To Delete The Current Record ? ", vbCritical + vbYesNo, "Data Deletion")
If res = vbYes Then
cnPatients.Execute ("delete from slip where ecode='" & Code & "'")
Call clear
ElseIf res = vbNo Then
MsgBox "Deletion Cancled", vbInformation
End If
End If
End Sub

Private Sub cmd_sal_modify_Click()
Dim res
rs.Open "Select * From Slip", cnPatients, adOpenDynamic, adLockOptimistic
Code.Enabled = False
If Code = "" Then
MsgBox "There Is No Current Record", vbInformation
Else
res = MsgBox("Do You Want To Modify The Current Record ? ", vbCritical + vbYesNo, "Data Modification")
If res = vbYes Then
cnPatients.Execute ("delete from slip where ecode='" & Code & "'")
rs.AddNew
Call upload
rs.Update
Call clear
ElseIf res = vbNo Then
MsgBox "Modifcation Cancled", vbInformation
End If
End If
rs.Close
End Sub

Private Sub cmd_sal_refresh_Click()

rs.Open "select * from slip where ecode='" & Code & "' ", con, adOpenDynamic, adLockOptimistic

If rs.RecordCount = 0 Then
MsgBox "No Current Records To Refresh", vbInformation, "Alert"

ElseIf rs.RecordCount > 0 Then
With rs
.Requery
.MoveFirst
End With
End If
rs.Close
End Sub

'##ModelId=3CBB305A019D
Private Sub cmd_sal_save_Click()
rs.Open "select * from slip", con, adOpenDynamic, adLockOptimistic
rs.AddNew
Call upload
rs.Update
Call clear
rs.Close
cmd_salary_add.Enabled = True
cmd_sal_save.Enabled = False
End Sub

Private Sub cmd_sal_search_Click()
Dim str_search_number
Code.Enabled = False
str_search_number = InputBox("Enter The Employee Number", "Data Search")
rs.Open "select * from slip", cnPatients, adOpenDynamic, adLockOptimistic
rs.MoveFirst
While Not rs.EOF
If rs.Fields(0) = str_search_number Then
MsgBox "Record Found"
Call download
End If
rs.MoveNext
Wend
Code.Enabled = False
nm.Enabled = False
add.Enabled = False
bp.Enabled = False
desig.Enabled = False
rs.Close
End Sub

'##ModelId=3CBB305A01A6
Private Sub cmd_sal_viewall_Click()
frm_salary_view.Show
End Sub

'##ModelId=3CBB305A01A7
Private Sub cmd_salary_add_Click()
cmd_salary_add.Enabled = False
cmd_sal_save.Enabled = True
Call clear
End Sub

'##ModelId=3CBB305A01B0
Private Sub cmd_salary_back_Click()
Unload Me
frm_employee.Show
End Sub

'##ModelId=3CBB305A01B1
Private Sub Code_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    nm.SetFocus
End If
End Sub
'##ModelId=3CBB305A01BA
Private Sub Code_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub

'##ModelId=3CBB305A01BC
Private Sub Command1_Click()

End Sub

'##ModelId=3CBB305A01C4
Private Sub da_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    add.SetFocus
End If
End Sub

'##ModelId=3CBB305A01C7
Private Sub da_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub
'##ModelId=3CBB305A01CF
Private Sub da_LostFocus()
    gpa.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text)
    Net.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) - Val(pf.Text) - Val(ins.Text) - Val(it.Text) - Val(pt.Text)
End Sub
'##ModelId=3CBB305A01D0
Private Sub deduct_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub
'##ModelId=3CBB305A01D9
Private Sub deduct_LostFocus()
    deduct.Text = Val(pf.Text) + Val(ins.Text) + Val(it.Text) + Val(pt.Text)
    Net.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) - Val(pf.Text) - Val(ins.Text) - Val(it.Text) - Val(pt.Text)
End Sub
'##ModelId=3CBB305A01DA
Private Sub Desig_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    bp.SetFocus
End If
End Sub
'##ModelId=3CBB305A01E3
Private Sub Desig_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 91 Or KeyAscii >= 97 And KeyAscii <= 123 Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 13 Then
    If Len(nm.Text) = 0 Or KeyAscii = 32 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 32 Then
          ch = "True"
        End If
    Else
        If ch = "true" Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            ch = "false"
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
    Else
        MsgBox ("Invalid Input")
        KeyAscii = 0
    End If
End Sub

'##ModelId=3CBB305A01E5
Private Sub Form_Load()

Set con = New ADODB.Connection
Set rs = New ADODB.Recordset
'con.CursorLocation = adUseClient
'con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\pay\Pay.mdb;persist security info=false"

cmd_salary_add.Enabled = True
cmd_sal_save.Enabled = False

Call clear

End Sub



'##ModelId=3CBB305A01ED
Private Sub gpa_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub
'##ModelId=3CBB305A01EF
Private Sub gpa_LostFocus()
    gpa.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text)
    Net.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) - Val(pf.Text) - Val(ins.Text) - Val(it.Text) - Val(pt.Text)
End Sub
'##ModelId=3CBB305A01F7
Private Sub hr_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
ca.SetFocus
End If
End Sub
'##ModelId=3CBB305A0201
Private Sub hr_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub
'##ModelId=3CBB305A0203
Private Sub hr_LostFocus()
    gpa.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text)
    Net.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) - Val(pf.Text) - Val(ins.Text) - Val(it.Text) - Val(pt.Text)
End Sub
'##ModelId=3CBB305A020B
Private Sub ins_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  it.SetFocus
End If
End Sub
'##ModelId=3CBB305A020E
Private Sub ins_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub
'##ModelId=3CBB305A0215
Private Sub ins_LostFocus()
    deduct.Text = Val(pf.Text) + Val(ins.Text) + Val(it.Text) + Val(pt.Text)
    Net.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) - Val(pf.Text) - Val(ins.Text) - Val(it.Text) - Val(pt.Text)
End Sub
'##ModelId=3CBB305A0216
Private Sub it_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  pt.SetFocus
End If
End Sub
'##ModelId=3CBB305A021F
Private Sub it_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub
'##ModelId=3CBB305A0221
Private Sub it_LostFocus()
    deduct.Text = Val(pf.Text) + Val(ins.Text) + Val(it.Text) + Val(pt.Text)
    Net.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) - Val(pf.Text) - Val(ins.Text) - Val(it.Text) - Val(pt.Text)
End Sub

'##ModelId=3CBB305A0229
Private Sub Net_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub
'##ModelId=3CBB305A022B
Private Sub Net_LostFocus()
    gpa.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text)
    deduct.Text = Val(pf.Text) + Val(ins.Text) + Val(it.Text) + Val(pt.Text)
    Net.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) - Val(pf.Text) - Val(ins.Text) - Val(it.Text) - Val(pt.Text)
End Sub
'##ModelId=3CBB305A0233
Private Sub nm_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  add.SetFocus
End If
End Sub
'##ModelId=3CBB305A023D
Private Sub nm_KeyPress(KeyAscii As Integer)
If KeyAscii >= 65 And KeyAscii <= 91 Or KeyAscii >= 97 And KeyAscii <= 123 Or KeyAscii = 8 Or KeyAscii = 32 Or KeyAscii = 13 Then
    If Len(nm.Text) = 0 Or KeyAscii = 32 Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii = 32 Then
          ch = "True"
        End If
    Else
        If ch = "true" Then
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            ch = "false"
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
    End If
    Else
        MsgBox ("Invalid Input")
        KeyAscii = 0
    End If
End Sub
'##ModelId=3CBB305A023F
Private Sub pf_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  ins.SetFocus
End If
End Sub
'##ModelId=3CBB305A0247
Private Sub pf_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub
'##ModelId=3CBB305A0249
Private Sub pf_LostFocus()
    deduct.Text = Val(pf.Text) + Val(ins.Text) + Val(it.Text) + Val(pt.Text)
    Net.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) - Val(pf.Text) - Val(ins.Text) - Val(it.Text) - Val(pt.Text)
End Sub
'##ModelId=3CBB305A0251
Private Sub pt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  cmd_salary_add.SetFocus
End If
End Sub
'##ModelId=3CBB305A025B
Private Sub pt_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub
'##ModelId=3CBB305A025D
Private Sub pt_LostFocus()
    deduct.Text = Val(pf.Text) + Val(ins.Text) + Val(it.Text) + Val(pt.Text)
    Net.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) - Val(pf.Text) - Val(ins.Text) - Val(it.Text) - Val(pt.Text)
End Sub

'##ModelId=3CBB305A0265
Private Sub ta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  pf.SetFocus
End If
End Sub
'##ModelId=3CBB305A0268
Private Sub ta_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 8 Then
    KeyAscii = KeyAscii
Else
    KeyAscii = 0
End If
End Sub
'##ModelId=3CBB305A026F
Private Sub ta_LostFocus()
    gpa.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text)
    Net.Text = Val(bp.Text) + Val(da.Text) + Val(hr.Text) + Val(ca.Text) + Val(ta.Text) - Val(pf.Text) - Val(ins.Text) - Val(it.Text) - Val(pt.Text)
End Sub


Public Sub clear()
 Me.Code.Text = ""
    Me.nm.Text = ""
    Me.add.Text = ""
    Me.desig.Text = ""
    Me.bp.Text = ""
    Me.da.Text = ""
    Me.hr.Text = ""
    Me.ca.Text = ""
    Me.ta.Text = ""
    Me.pf.Text = ""
    Me.ins.Text = ""
    Me.it.Text = ""
    Me.pt.Text = ""
    Me.gpa.Text = ""
    Me.deduct.Text = ""
    Me.Net.Text = ""
End Sub

Public Sub upload()
    rs!ecode = Me.Code.Text
    rs!name = Me.nm.Text
    rs!dep = Me.desig.Text
    rs!address = Me.add.Text
    rs!basic = Me.bp.Text
    rs!da = Me.da.Text
    rs!hra = Me.hr.Text
    rs!cca = Me.ca.Text
    rs!trans = Me.ta.Text
    rs!gpf = Me.pf.Text
    rs!ins = Me.ins.Text
    rs!itax = Me.it.Text
    rs!ptax = Me.pt.Text
    rs!gross = Me.gpa.Text
    rs!deduct = Me.deduct.Text
    rs!Net = Me.Net.Text
End Sub


Public Sub download()
With frm_add_salary_info
.Code.Text = rs!ecode
.nm.Text = rs!name
.desig.Text = rs!dep
.add.Text = rs!address
.bp.Text = rs!basic
.da.Text = rs!da
.hr.Text = rs!hra
.ca.Text = rs!cca
.ta.Text = rs!trans
.pf.Text = rs!gpf
.ins.Text = rs!ins
.it.Text = rs!itax
.pt.Text = rs!ptax
.gpa.Text = rs!gross
.deduct.Text = rs!deduct
.Net.Text = rs!Net
End With
End Sub
