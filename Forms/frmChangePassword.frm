VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmChangePassword 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change User Password"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5940
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmChangePassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   5280
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      AutoCenterForm  =   -1  'True
      FormHeightDT    =   3645
      FormWidthDT     =   6030
      FormScaleHeightDT=   3165
      FormScaleWidthDT=   5940
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtReType 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "•"
      TabIndex        =   7
      Top             =   2040
      Width           =   2655
   End
   Begin VB.TextBox txtNewPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "•"
      TabIndex        =   6
      Top             =   1440
      Width           =   2655
   End
   Begin VB.TextBox txtOldPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "•"
      TabIndex        =   5
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtUserName 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Caption         =   "Retype New Password"
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
      Left            =   120
      TabIndex        =   3
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Caption         =   "New Password"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   "Old Password"
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
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "User Name"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
On Error Resume Next
If txtUserName = "" Then
    MsgBox "Please Enter the User Name", vbCritical
    txtUserName.SetFocus
    Exit Sub
End If

If txtNewPassword <> txtReType Then
    MsgBox "In correct retype password", vbCritical
    txtReType.SetFocus
    Exit Sub
End If

Dim rsChk As Recordset
Set rsChk = New ADODB.Recordset

rsChk.Open "Select * from User_Details where User_Name = '" & txtUserName & "'", cnPatients, adOpenDynamic, adLockPessimistic

If rsChk.EOF = True Then
    MsgBox "Invalid User Name", vbCritical
    txtUserName.SetFocus
    rsChk.Close
    Exit Sub
End If

Dim pass As String

If rsChk.EOF = False Then
    If rsChk.RecordCount = 1 Then
        pass = Functions.Decrypt(rsChk(1))
            If txtOldPassword <> pass Then
                MsgBox "Invalid Old Password", vbCritical
                rsChk.Close
                Exit Sub
            End If
            
        rsChk(1) = Functions.Encrypt(txtNewPassword)
        rsChk.Update
        MsgBox "Password has been changed sucessfully", vbInformation
        rsChk.Close
        Unload Me
    Else
     MsgBox "An Error Occured", vbCritical
     rsChk.Close
     Exit Sub
    End If
End If
End Sub
