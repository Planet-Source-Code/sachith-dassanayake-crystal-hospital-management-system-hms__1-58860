VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmAddDocAppointments 
   BackColor       =   &H00FF8080&
   Caption         =   "Add Doctor Appointments"
   ClientHeight    =   10725
   ClientLeft      =   3195
   ClientTop       =   1560
   ClientWidth     =   13140
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddDoctorAppointments.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10725
   ScaleWidth      =   13140
   WindowState     =   2  'Maximized
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   10920
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      AutoCenterForm  =   -1  'True
      FormHeightDT    =   11235
      FormWidthDT     =   13260
      FormScaleHeightDT=   10725
      FormScaleWidthDT=   13140
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Doctor Schedule"
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
      Height          =   3015
      Left            =   120
      TabIndex        =   18
      Top             =   7560
      Width           =   12855
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2295
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         AllowUserResizing=   3
         FormatString    =   "Doctor ID | Channeling Days | Time In | Time Out | Schedule ID | Notes"
      End
   End
   Begin VB.CommandButton cmdDocSched 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Doctor Schedule"
      Height          =   855
      Left            =   4680
      Picture         =   "frmAddDoctorAppointments.frx":57E2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Frame fra 
      BackColor       =   &H00FF8080&
      Height          =   6495
      Left            =   6720
      TabIndex        =   16
      Top             =   840
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton cmdPatSel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OK"
         Height          =   915
         Left            =   1680
         Picture         =   "frmAddDoctorAppointments.frx":593F
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         Height          =   915
         Left            =   3480
         Picture         =   "frmAddDoctorAppointments.frx":5E15
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CommandButton cmdDocIDSel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OK"
         Height          =   915
         Left            =   1680
         Picture         =   "frmAddDoctorAppointments.frx":6261
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   5280
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid fdata 
         Height          =   4095
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   7223
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   3
      End
   End
   Begin VB.Frame frameAppointment 
      BackColor       =   &H00FF8080&
      Caption         =   "Appointment Details"
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
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6495
      Begin VB.ComboBox cmbPatientID 
         Height          =   315
         Left            =   2040
         TabIndex        =   1
         Top             =   480
         Width           =   2415
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   915
         Left            =   3720
         Picture         =   "frmAddDoctorAppointments.frx":6737
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5280
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddApointment 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save Appointment"
         Height          =   915
         Left            =   1440
         Picture         =   "frmAddDoctorAppointments.frx":6C3B
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5280
         Width           =   1335
      End
      Begin VB.ComboBox cmbDoctorID 
         Height          =   315
         ItemData        =   "frmAddDoctorAppointments.frx":70D6
         Left            =   2040
         List            =   "frmAddDoctorAppointments.frx":70D8
         TabIndex        =   2
         Top             =   1200
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTPTime1 
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Top             =   4560
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   503
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
         CustomFormat    =   "hh:mm"
         Format          =   20709378
         CurrentDate     =   38329.0833333333
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   2295
         Left            =   2040
         TabIndex        =   4
         Top             =   1920
         Width           =   3855
         _Version        =   524288
         _ExtentX        =   6800
         _ExtentY        =   4048
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2004
         Month           =   12
         Day             =   25
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   7
         GridCellEffect  =   0
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPDate1 
         Height          =   285
         Left            =   2040
         TabIndex        =   11
         Top             =   1920
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   503
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
         Format          =   20709377
         CurrentDate     =   38331
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "Patient ID"
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
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF8080&
         Caption         =   "Doctor ID"
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
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FF8080&
         Caption         =   "Appoinmnet Date"
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
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FF8080&
         Caption         =   "Appoinment Time"
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
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   4560
         Width           =   1815
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "ADD DOCTOR APPOINTMENT"
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
      Left            =   4080
      TabIndex        =   20
      Top             =   240
      Width           =   5790
   End
End
Attribute VB_Name = "frmAddDocAppointments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flex As Integer
Dim co, del_i  As Integer
Dim dupid As String
Dim rgrid, COUNT1, rs As Integer
Dim ch, ch1 As Integer




Private Sub Calendar1_Click()
DTPDate1.Value = Calendar1.Value
End Sub

Private Sub cmbDoctorID_Change()
cmdDocSched_Click


Dim rsDocPat As Recordset
Set rsDocPat = New ADODB.Recordset

flex = 20
fra.Visible = True
cmdPatSel.Visible = False
cmdDocIDSel.Visible = True

If ch = 0 Then

If Trim(cmbDoctorID.Text) = "" Or Len(cmbDoctorID.Text) = 0 Then
    rsDocPat.Open "select * from Doctor_Details", cnPatients, adOpenDynamic, adLockPessimistic
Else
    rsDocPat.Open "select * from Doctor_Details where Doctor_ID like '" & UCase(Trim(cmbDoctorID.Text)) & "%'", cnPatients, adOpenDynamic, adLockPessimistic
End If

If rsDocPat.RecordCount > 0 Then
    rsDocPat.MoveFirst
    fdata.clear
    fdata.FixedRows = 1
    fdata.Rows = 2
    fdata.FormatString = "Doctor ID" & vbTab & "First Name" & vbTab & "Last Name" & vbTab & "Specialization" & vbTab & "Qualification"
   fdata.ColWidth(0) = 1500
fdata.ColWidth(1) = 1800
fdata.ColWidth(2) = 1800
fdata.ColWidth(3) = 1500
fdata.ColWidth(4) = 2000
    fdata.TextMatrix(1, 0) = rsDocPat.Fields(0)
    fdata.TextMatrix(1, 1) = rsDocPat.Fields(1)
    fdata.TextMatrix(1, 2) = rsDocPat.Fields(2)
    fdata.TextMatrix(1, 3) = rsDocPat.Fields(9)
    fdata.TextMatrix(1, 4) = rsDocPat.Fields(10)
    rsDocPat.MoveNext
    
    While Not rsDocPat.EOF
        fdata.AddItem rsDocPat.Fields(0) & vbTab & rsDocPat.Fields(1) & vbTab & rsDocPat.Fields(2) & vbTab & rsDocPat.Fields(9) & vbTab & rsDocPat.Fields(10)
        rsDocPat.MoveNext
    Wend
Else
MsgBox "Name/ID Doesn't Exist", vbCritical + vbOKOnly, "Invalid Name/ID"
fdata.clear
fdata.FixedRows = 1
fdata.Rows = 2
fdata.FormatString = "Doctor ID" & vbTab & "First Name" & vbTab & "Last Name" & vbTab & "Specialization" & vbTab & "Qualification"
fdata.ColWidth(0) = 1500
fdata.ColWidth(1) = 1800
fdata.ColWidth(2) = 1800
fdata.ColWidth(3) = 1500
fdata.ColWidth(4) = 2000
End If
rsDocPat.Close
End If

End Sub

Private Sub cmbDoctorID_Click()
cmbDoctorID_Change
cmdDocSched_Click
End Sub

Private Sub cmbDoctorID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  fdata.SetFocus
  frameAppointment.Enabled = False
End If
End Sub

Private Sub cmbPatientID_Change()


Dim rsDocPat As Recordset
Set rsDocPat = New ADODB.Recordset

flex = 10
fra.Visible = True
cmdPatSel.Visible = True
cmdDocIDSel.Visible = False



If ch = 0 Then

If Trim(cmbPatientID.Text) = "" Or Len(cmbPatientID.Text) = 0 Then
    rsDocPat.Open "select * from Patient_Details", cnPatients, adOpenDynamic, adLockPessimistic
Else
    rsDocPat.Open "select * from Patient_Details where Patient_ID like '" & UCase(Trim(cmbPatientID.Text)) & "%'", cnPatients, adOpenDynamic, adLockPessimistic
End If

If rsDocPat.RecordCount > 0 Then
    rsDocPat.MoveFirst
    fdata.clear
    fdata.FixedRows = 1
    fdata.Rows = 2
    fdata.FormatString = "Patient ID" & vbTab & "First Name" & vbTab & "Last Name" & vbTab & "Gender" & vbTab & "Address"
    fdata.ColWidth(0) = 1500
    fdata.ColWidth(1) = 1800
    fdata.ColWidth(2) = 1800
    fdata.ColWidth(3) = 1600
    fdata.ColWidth(4) = 1600
    fdata.TextMatrix(1, 0) = rsDocPat.Fields(0)
    fdata.TextMatrix(1, 1) = rsDocPat.Fields(1)
    fdata.TextMatrix(1, 2) = rsDocPat.Fields(2)
    fdata.TextMatrix(1, 3) = rsDocPat.Fields(3)
    fdata.TextMatrix(1, 4) = rsDocPat.Fields(4)
    rsDocPat.MoveNext
    
    While Not rsDocPat.EOF
        fdata.AddItem rsDocPat.Fields(0) & vbTab & rsDocPat.Fields(1) & vbTab & rsDocPat.Fields(2) & vbTab & rsDocPat.Fields(3) & vbTab & rsDocPat.Fields(4)
        rsDocPat.MoveNext
    Wend
Else
MsgBox "Patient ID Doesn't Exist", vbCritical + vbOKOnly, "Invalid Name/ID"
fdata.clear
fdata.FixedRows = 1
fdata.Rows = 2
fdata.FormatString = "Patient ID" & vbTab & "First Name" & vbTab & "Last Name" & vbTab & "Gender" & vbTab & "Address"
fdata.ColWidth(0) = 1500
fdata.ColWidth(1) = 1800
fdata.ColWidth(2) = 1800
fdata.ColWidth(3) = 1000
fdata.ColWidth(4) = 1100
End If
rsDocPat.Close
End If



End Sub

Private Sub cmbPatientID_Click()
cmbPatientID_Change
End Sub

Private Sub cmbPatientID_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
  fdata.SetFocus
  frameAppointment.Enabled = False
End If








End Sub

Private Sub cmdAddApointment_Click()
'On Error GoTo AddErr

Dim rsAddBill As Recordset
Dim rsAddAppointment As Recordset
Dim rsInfo As New Recordset
Dim rsSched As Recordset
Dim rsPrevApp As Recordset

Dim StrDoctorID As String
Dim StrServiceID As String
Dim PID As String
Dim strTime As String
Dim BID As String
Dim strAvailDays As String
Dim strDocIn As String
Dim strDocOut As String
Dim strDate As String
Dim strAll As String
Dim strDIn() As String
Dim strDOut() As String
Dim arrDays() As String

Dim ID As Boolean
Dim flag As Boolean
Dim proceed As Boolean
Dim AppPos As Boolean

Dim i As Integer
Dim NoOfSchedules As Integer
Dim fl As Integer

Dim newTime As Double
Dim newTime2 As Double
Dim strInTime As Date
Dim strOutTime As Date
Dim lastAppTime As Date




Set rsAddAppointment = New ADODB.Recordset
Set rsInfo = New ADODB.Recordset
Set rsSched = New ADODB.Recordset
Set rsAddBill = New ADODB.Recordset
Set rsPrevApp = New ADODB.Recordset





 If cnPatients.State = adStateOpen Then
    'cnPatients.BeginTrans
    
    
        ' Adding data to the Appointment_Bill table
        strTime = DTPTime1.Value
        
                
        BID = Functions.UID(6, "OPBID_")    'Generate random Bill ID
        rsAddBill.Open "Select * from Appointment_Bill", cnPatients, adOpenKeyset, adLockPessimistic
        While rsAddBill.EOF = False
            If rsAddBill(0) = BID Then  ' If Bill ID found Generate Another Bill ID
                BID = Functions.UID(6, "OPBID_")
                rsAddBill.MoveFirst
            End If
            rsAddBill.MoveNext
        Wend
        
       
           
            StrDoctorID = cmbDoctorID.Text
                          
            ' Doctor Channeling Charges
            rsInfo.Open "Select * from Doctor_Details where Doctor_ID='" & StrDoctorID & "'", cnPatients, adOpenKeyset, adLockPessimistic
            If rsInfo.EOF = False Then
                strAmount = rsInfo![Doctor_CCharge]
            End If
            
                    ' Assigning temporary Values
                    HospitalCharge = 200
                    Discount = 20
                    GrandTotal = HospitalCharge + strAmount
                    NetValue = GrandTotal - Discount
                    AppBillID = BID
                    BillPatientID = cmbPatientID
        
                    ' End of BillData
            
            
            PID = Functions.UID(6, "DApp_") 'Generate Random Appointment ID
            'Generate Random Appointment ID
            rsAddAppointment.Open "Select * from Doctor_Appointment", cnPatients, adOpenKeyset, adLockPessimistic
            
            While rsAddAppointment.EOF = False
                If rsAddAppointment(0) = PID Then
                    PID = Functions.UID(6, "DApp_")
                    rsAddAppointment.MoveFirst
                End If
                rsAddAppointment.MoveNext
            Wend

            
             
            
           ' Possible Data Validation (If data is invalid it will exit the sub)
             
           ' If the Appointment date is less than the current date
            If cmbPatientID = "" Then
                MsgBox "Please enter a valid patient ID", vbCritical, "Out Patient Details"
            rsAddBill.Close
            rsAddAppointment.Close
            rsInfo.Close
            
                Exit Sub
            End If
            If cmbDoctorID = "" Then
                MsgBox "Please enter a valid Doctor ID", vbCritical, "Out Patient Details"
                Exit Sub
            End If
            
            
            If DTPDate1.Value < Date Then
                MsgBox "Appointment Date Should Be Greater Than Current Date", vbCritical, "Invalid Date"
                rsAddBill.Close
                rsAddAppointment.Close
                rsInfo.Close
           
                Exit Sub
            End If
                       
            If DTPDate1.Value = Date And strTime < Time Then
                MsgBox "Appointment Time Should Be Greater Than Current Time", vbCritical, "Invalid Date"
                rsAddBill.Close
                rsAddAppointment.Close
                rsInfo.Close
                Exit Sub
            End If
             
         
            
            rsSched.Open "Select * from Doctor_Schedule_Details where Doctor_ID='" & StrDoctorID & "'", cnPatients, adOpenDynamic, adLockPessimistic
            
            ' Retreive Doctor Available days from the table
            While rsSched.EOF = False
                strAvailDays = strAvailDays & rsSched![Doctor_AvaiDate] & "..."
                strDocIn = strDocIn & rsSched![Doctor_In] & "..."
                strDocOut = strDocOut & rsSched![Doctor_Out] & "..."
                rsSched.MoveNext
            Wend
            
            arrDays() = Split(strAvailDays, "...")
            strDIn() = Split(strDocIn, "...")
            strDOut() = Split(strDocOut, "...")
            
                   
            
            strDate = Left(Format(DTPDate1.Value, "dddd"), 3)
                       
            For i = 0 To UBound(arrDays)
                If InStr(1, arrDays(i), Left(Format(DTPDate1.Value, "dddd"), 3)) > 0 Then
                    AppPos = True ' Doctor is available on the selected date
                End If
            Next i
            
   

            fl = 1
            
            
            Debug.Print strTime
            For i = 0 To UBound(strDIn)
                Debug.Print "In : " & strDIn(i) & "    " & "Out : " & strDOut(i)
            Next
            

            For i = 0 To UBound(strDIn)
                If InStr(1, arrDays(i), strDate) > 0 Then
                    If strTime >= strDIn(i) And strTime <= strDOut(i) Then
                        strInTime = Format(strDIn(i), "short time")
                        strOutTime = Format(strDOut(i), "short time")
                        newTime = DateDiff("s", strDIn(i), strDOut(i)) / 60
                        newTime2 = newTime / 60
                        
                        Debug.Print "Total Time in Minutes : " & newTime & " Minutes"
                        Debug.Print "Total Time in hours   : " & newTime2 & " Hours"
                        Debug.Print "Possible Appointments : " & newTime / 15
                        Debug.Print "Docotor In            : " & strInTime
                        Debug.Print "Doctor Out            : " & strOutTime
                        fl = 0
                        Exit For
                    Else
                        fl = 1
                    End If
                End If
                
            Next i
            
                       
            
            If fl = 0 Then
                Debug.Print "Appointment Possible (Time and Date)"
            Else
                MsgBox "The Appointment Date or Time does not valid", vbInformation, "Out Patient Apointment"
            rsAddBill.Close
            rsAddAppointment.Close
            rsInfo.Close
            rsSched.Close
                Exit Sub
            End If

            
            rsPrevApp.Open "select * from Doctor_Appointment where Doctor_ID='" & StrDoctorID & "' and  Appointment_Date=#" & DTPDate1.Value & "# and Appointment_Time >= #" & strInTime & "# and  Appointment_time <= #" & strOutTime & "#", cnPatients, adOpenKeyset, adLockPessimistic
            NoOfSchedules = rsPrevApp.RecordCount
            
            While rsPrevApp.EOF = False
                Debug.Print "Current Appointments at :" & rsPrevApp![Appointment_Time]
                lastAppTime = Format(rsPrevApp![Appointment_Time], "short time")
                rsPrevApp.MoveNext
            Wend
            

            
            If NoOfSchedules > 0 Then
            If lastAppTime >= Format(DTPTime1.Value, "short time") Then
                MsgBox "Appointment Time should be greater than previous appointment"
                rsAddBill.Close
                rsAddAppointment.Close
                rsInfo.Close
                rsPrevApp.Close
                rsSched.Close
                Exit Sub
            End If
            End If
            
            
            If NoOfSchedules >= (newTime / 15) Then      '15 mins per patient
                MsgBox "No Space Available For This Appointment", vbInformation, "Doctor Appointments"
                rsAddBill.Close
                rsAddAppointment.Close
                rsInfo.Close
                rsPrevApp.Close
                rsSched.Close
                Exit Sub
            End If
            
            
                 
            ' Add data to the database if no errors occured
      
           
            If MsgBox("Are you sure you want to add this record to the database?", vbYesNo, "Add Doctor Appoinment") = vbYes Then
                cnPatients.Execute ("Insert into Doctor_Appointment values('" & PID & "','" & cmbPatientID & "','" & StrDoctorID & "','" & DTPDate1.Value & "','" & strTime & "')")
                cnPatients.Execute ("Insert into Appointment_Bill values('" & BID & "','" & PID & "','" & cmbPatientID & "','" & Format(Date, "mm/dd/yy") & "'," & strAmount & "," & HospitalCharge & "," & GrandTotal & "," & Discount & "," & Val(NetValue) & ")")
            
            'cnPatients.CommitTrans
            rsAddBill.Close
            rsAddAppointment.Close
            rsInfo.Close
            rsPrevApp.Close
            rsSched.Close
            
            Unload Me
            frmAppoinmnetCharges.Show
            
            Else
            Exit Sub
            End If
            
      
    Else
        'when database connection error occurs
        MsgBox "Database Connection Error", vbCritical, "SD Hospitals PVT LTD"
    End If
    Exit Sub
AddErr:
MsgBox Err.Description
End Sub

Private Sub cmdCancel_Click()
fra.Visible = False
frameAppointment.Enabled = True
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDocIDSel_Click()

Dim rsDocPat As Recordset
Set rsDocPat = New ADODB.Recordset

ch = 1

rsDocPat.Open "select * from Doctor_Details where Doctor_ID like '" & UCase(Trim(cmbDoctorID.Text)) & "%'", cnPatients, adOpenDynamic, adLockPessimistic
If rsDocPat.RecordCount > 0 Then
dupid = fdata.TextMatrix(fdata.Row, 0)
cmbDoctorID.Text = dupid
'cmbDoctorID.SetFocus

Else
MsgBox "Select the Appropiate Data from the [[GridBox]]", vbInformation + vbOKOnly, "Error"

End If

rsDocPat.Close
dupid = 0
ch = 0
fra.Visible = False
frameAppointment.Enabled = True





End Sub

Private Sub cmdPatientID_Click()
frmSelPatientID.Show
End Sub

Private Sub cmdDocSched_Click()

Dim rsSched As Recordset
Set rsSched = New ADODB.Recordset


rsSched.Open "Select * from Doctor_Schedule_Details where Doctor_ID='" & cmbDoctorID & "'", cnPatients, adOpenKeyset, adLockPessimistic
            
            ' Retreive Doctor Available days from the table
If rsSched.EOF = False Then
        

    
    rsSched.MoveFirst
    
    MSFlexGrid1.clear
    
    With MSFlexGrid1
        .clear
        .Rows = 1
        .Cols = rsSched.Fields.Count
   
  

        While Not rsSched.EOF
            .Rows = .Rows + 1
         .Row = .Rows - 1


        .TextMatrix(.Row, 0) = rsSched(1)
        .TextMatrix(.Row, 1) = rsSched(4)
        .TextMatrix(.Row, 2) = rsSched(2)
        .TextMatrix(.Row, 3) = rsSched(3)
        .TextMatrix(.Row, 4) = rsSched(0)
        .TextMatrix(.Row, 5) = rsSched(5)
            
            
            
            
            rsSched.MoveNext
        Wend
    
    
        .TextMatrix(0, 0) = "Doctor ID"
        .TextMatrix(0, 1) = "Available Date"
        .TextMatrix(0, 2) = "Time In"
        .TextMatrix(0, 3) = "Time Out"
        .TextMatrix(0, 4) = "Schedule ID"
        .TextMatrix(0, 5) = "Notes"
   
        .FixedRows = 1
        .RowHeight(0) = .RowHeight(1) * 1.5
        'Functions.SizeColumns MSFlexGrid1, Me
        Functions.SizeColumnHeaders MSFlexGrid1, Me
    End With
    
 Else
 Debug.Print "No Records Found"
 End If
    
End Sub

Private Sub cmdPatSel_Click()
Dim rsDocPat As Recordset
Set rsDocPat = New ADODB.Recordset

ch = 1

rsDocPat.Open "select * from Patient_Details where Patient_ID like '" & UCase(Trim(cmbPatientID.Text)) & "%'", cnPatients, adOpenDynamic, adLockPessimistic
If rsDocPat.RecordCount > 0 Then
dupid = fdata.TextMatrix(fdata.Row, 0)
cmbPatientID.Text = dupid
frameAppointment.Enabled = True
cmbDoctorID.SetFocus

Else
MsgBox "Select the Appropiate Data from the [[GridBox]]", vbInformation + vbOKOnly, "Error"

End If

rsDocPat.Close
dupid = 0
ch = 0
fra.Visible = False
frameAppointment.Enabled = True

End Sub


Private Sub fdata_Click()

dupid = fdata.TextMatrix(fdata.Row, 0)

End Sub

Private Sub fdata_KeyPress(KeyAscii As Integer)

Dim rsDocPat As Recordset
Set rsDocPat = New ADODB.Recordset

If KeyAscii = 13 Then
    If flex = 20 Then
          dupid = fdata.TextMatrix(fdata.Row, 0)
        dupid1 = fdata.TextMatrix(fdata.Row, 0)
        customer_code = fdata.TextMatrix(fdata.Row, 0)
        ch = 1

        rsDocPat.Open "select * from Doctor_Details where Doctor_ID like '" & cmbDoctorID.Text & "%'", cnPatients, adOpenDynamic, adLockPessimistic
        If rsDocPat.RecordCount > 0 Then
            dupid = fdata.TextMatrix(fdata.Row, 0)
            cmbDoctorID.Text = dupid
           
            
        Else
            MsgBox "Select the Appropiate Data from the [[GridBox]]", vbInformation + vbOKOnly, "Error"
        End If
            rsDocPat.Close
            dupid = 0
            ch = 0
            fra.Visible = False
            frameAppointment.Enabled = True
    ElseIf flex = 10 Then
        dupid = fdata.TextMatrix(fdata.Row, 0)
        dupid1 = fdata.TextMatrix(fdata.Row, 0)
        customer_code = fdata.TextMatrix(fdata.Row, 0)
        ch = 1

        rsDocPat.Open "select * from Patient_Details where Patient_ID like '" & cmbPatientID.Text & "%'", cnPatients, adOpenDynamic, adLockPessimistic
        If rsDocPat.RecordCount > 0 Then
            dupid = fdata.TextMatrix(fdata.Row, 0)
            cmbPatientID.Text = dupid
            Debug.Print dupid
            frameAppointment.Enabled = True
            cmbDoctorID.SetFocus
           
        Else
            MsgBox "Select the Appropiate Data from the [[GridBox]]", vbInformation + vbOKOnly, "Error"
        End If
            rsDocPat.Close
            dupid = 0
            ch = 0
            fra.Visible = False
            frameAppointment.Enabled = True
    End If
    
ElseIf KeyAscii = 27 Then
    If flex = 20 Then
        fra.Visible = False
        cmbDoctorID.SetFocus
        frameAppointment.Enabled = True
    Else
        fra.Visible = False
        cmbPatientID.SetFocus
        frameAppointment.Enabled = True
    End If
End If





















End Sub

Private Sub Form_Activate()
Call Functions.DisableMenu
End Sub

Private Sub Form_Load()

Call Functions.DisableMenu

Dim SQL1 As String
Dim rsDoctors As Recordset
Dim rsPID As Recordset

Set rsDoctors = New ADODB.Recordset
Set rsPID = New ADODB.Recordset

COUNT1 = 0
rs = 1
del_i = 0
rgrid = 0
ch = 0
ch1 = 0


If cnPatients.State = adStateOpen Then
    'create sql statements
    SQL1 = "SELECT * FROM Doctor_Details"
    
    rsDoctors.Open SQL1, cnPatients, adOpenStatic, adLockPessimistic
       
    
    ' Add ID's to Combo Box
    While rsDoctors.EOF = False
        cmbDoctorID.AddItem rsDoctors(0)
        rsDoctors.MoveNext
    Wend
   
   
    
    rsPID.Open "select * from Patient_Details", cnPatients, adOpenStatic, adLockPessimistic
    
    While rsPID.EOF = False
        cmbPatientID.AddItem rsPID(0)
        rsPID.MoveNext
    Wend
  

   
   
    rsDoctors.Close
    rsPID.Close
    
    
    DTPDate1.Value = Date
    DTPTime1.Value = Time
    Calendar1.Value = Date
    
    MSFlexGrid1.ColWidth(0) = 1500
    MSFlexGrid1.ColWidth(1) = 3500
    MSFlexGrid1.ColWidth(2) = 2000
    MSFlexGrid1.ColWidth(3) = 2000
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.ColWidth(5) = 2000
    
    fdata.RowHeight(0) = fdata.RowHeight(1) * 1.5
    


Else
    'when database connection error occurs
    MsgBox "Database Connection Error", vbCritical, "SD Hospitals PVT LTD"
End

End If

End Sub



Private Sub Command2_Click()
frmDoctorDetails.Show
End Sub




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Functions.EnableMenu
End Sub
