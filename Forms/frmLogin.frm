VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2265
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5820
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Height          =   390
      Left            =   1800
      TabIndex        =   4
      Top             =   1680
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3360
      TabIndex        =   5
      Top             =   1680
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3240
      PasswordChar    =   "•"
      TabIndex        =   3
      Top             =   990
      Width           =   2325
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   240
      Picture         =   "frmLogin.frx":0E42
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000005&
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   1905
      TabIndex        =   0
      Top             =   495
      Width           =   1200
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H80000005&
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   1905
      TabIndex        =   2
      Top             =   1005
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean
Dim Counter As Integer

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Unload MDIMain
    Unload Me
    
End Sub

Private Sub cmdOK_Click()
On Error Resume Next
Dim TempPass As String
If Counter = 3 Then
    MsgBox "You have exceeded the number of attempts"
    cmdCancel_Click
    Exit Sub
End If
    
    'check for correct password
      
    
    Dim rsCheckUser As Recordset
    Set rsCheckUser = New ADODB.Recordset
    
    Dim rsAddLog As Recordset
    Set rsAddLog = New ADODB.Recordset
    Dim logID As String
    
    
rsCheckUser.Open "Select * from User_Details where User_Name = '" & txtUserName & "'", cnPatients, adOpenDynamic, adLockReadOnly
     
If rsCheckUser.EOF = False Then
TempPass = Functions.Decrypt(rsCheckUser(1))
    If txtPassword = TempPass Then
        UserCategory = rsCheckUser(2)
        User = txtUserName
        LogDate = Format(Date, "Short Date")
        LogTime = Format(Time, "Short Time")
        

        LoginSucceeded = True
       
       
       
        logID = Functions.UID(6, "LogID_")
        rsAddLog.Open "Select * from Log_Details", cnPatients, adOpenDynamic, adLockPessimistic
             
            While rsAddLog.EOF = False
                If rsAddLog(0) = logID Then
                    logID = Functions.UID(6, "LogID_")
                    rsAddLog.MoveFirst
                Else

                End If
                rsAddLog.MoveNext
            Wend
            
            rsAddLog.AddNew
            rsAddLog(0) = logID
            rsAddLog(1) = txtUserName
            rsAddLog(2) = Format(Date, "short date")
            rsAddLog(3) = Format(Time, "short time")
            rsAddLog.Update
             
            rsAddLog.Close
        Call Privilages
        Load MDIMain
        Load frmTip

        Unload frmLogin
        
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
        Counter = Counter + 1
    End If
Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
        Counter = Counter + 1
End If

    rsCheckUser.Close
End Sub


