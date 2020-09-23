VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmAddSerAppointments 
   BackColor       =   &H00FF8080&
   Caption         =   "Add Hospital Service Appointments"
   ClientHeight    =   11685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13740
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddSerAppointments.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11685
   ScaleWidth      =   13740
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Height          =   5175
      Left            =   360
      TabIndex        =   9
      Top             =   1320
      Width           =   6015
      Begin VB.CommandButton cmdPatientID 
         Caption         =   "..."
         Height          =   315
         Left            =   4320
         TabIndex        =   16
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtPatientID 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   15
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox cmbHospitalServiceID 
         Height          =   315
         Left            =   2040
         TabIndex        =   14
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   315
         Left            =   4320
         TabIndex        =   13
         Top             =   960
         Width           =   495
      End
      Begin VB.ComboBox cmbPatientID 
         Height          =   315
         ItemData        =   "frmAddSerAppointments.frx":57E2
         Left            =   2040
         List            =   "frmAddSerAppointments.frx":57E4
         TabIndex        =   12
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdSerSched 
         Caption         =   "Services Schedule"
         Height          =   375
         Left            =   5160
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   2295
         Left            =   2040
         TabIndex        =   11
         Top             =   1560
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
      Begin MSComCtl2.DTPicker DTPTime1 
         Height          =   285
         Left            =   2040
         TabIndex        =   17
         Top             =   4440
         Width           =   2175
         _ExtentX        =   3836
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
         Format          =   45875202
         CurrentDate     =   38329.0833333333
      End
      Begin MSComCtl2.DTPicker DTPDate1 
         Height          =   285
         Left            =   2040
         TabIndex        =   18
         Top             =   1560
         Width           =   2175
         _ExtentX        =   3836
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
         Format          =   45875201
         CurrentDate     =   38331
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "Patient ID"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF8080&
         Caption         =   "Hospital Service ID"
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FF8080&
         Caption         =   "Appoinmnet Date"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FF8080&
         Caption         =   "Appoinment Time"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   4440
         Width           =   1695
      End
   End
   Begin VB.Frame fra 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   6480
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CommandButton cmdDocIDSel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2280
         Picture         =   "frmAddSerAppointments.frx":57E6
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   4080
         Picture         =   "frmAddSerAppointments.frx":5CBC
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4080
         Width           =   1215
      End
      Begin VB.CommandButton cmdPatSel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2280
         Picture         =   "frmAddSerAppointments.frx":6108
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4080
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid fdata 
         Height          =   3375
         Left            =   600
         TabIndex        =   8
         Top             =   480
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   5953
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   3
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Service Schedule"
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   600
      TabIndex        =   2
      Top             =   6960
      Width           =   12495
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2295
         Left            =   1320
         TabIndex        =   3
         Top             =   480
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         AllowUserResizing=   3
         FormatString    =   "Service ID | Channeling Days | Time In | Time Out | Schedule ID | Notes"
      End
   End
   Begin VB.CommandButton cmdAddApointment 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Save Appointment"
      Height          =   855
      Left            =   4560
      Picture         =   "frmAddSerAppointments.frx":65DE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10200
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   855
      Left            =   6960
      Picture         =   "frmAddSerAppointments.frx":6A79
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   10200
      Width           =   1935
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      FormHeightDT    =   12195
      FormWidthDT     =   13860
      FormScaleHeightDT=   11685
      FormScaleWidthDT=   13740
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "HOSPITAL SERVICE APPOINTMENT"
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
      Left            =   3600
      TabIndex        =   23
      Top             =   360
      Width           =   7035
   End
End
Attribute VB_Name = "frmAddSerAppointments"
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

Private Sub cmbHospitalServiceID_Change()
fdata.RowSel = 1
cmdSerSched_Click


Dim rsSerPat As Recordset
Set rsSerPat = New ADODB.Recordset

flex = 20
fra.Visible = True
cmdPatSel.Visible = False
cmdDocIDSel.Visible = True

If ch = 0 Then

If Trim(cmbHospitalServiceID) = "" Or Len(cmbHospitalServiceID) = 0 Then
    rsSerPat.Open "select * from Services", cnPatients, adOpenDynamic, adLockPessimistic
Else
    rsSerPat.Open "select * from Services where Channel_Service_ID like '" & UCase(Trim(cmbHospitalServiceID.Text)) & "%'", cnPatients, adOpenDynamic, adLockPessimistic
End If

If rsSerPat.RecordCount > 0 Then
    rsSerPat.MoveFirst
    fdata.clear
    fdata.FixedRows = 1
    fdata.Rows = 2
    fdata.FormatString = "Service ID" & vbTab & "Service Name" & vbTab & "Duration" & vbTab & "Charge" & vbTab & "Notes"
    fdata.ColWidth(0) = 1500
    fdata.ColWidth(1) = 1800
    fdata.ColWidth(2) = 2000
    fdata.ColWidth(3) = 2000
    'fdata.ColWidth(4) = 2000
    fdata.TextMatrix(1, 0) = rsSerPat.Fields(0)
    fdata.TextMatrix(1, 1) = rsSerPat.Fields(1)
    fdata.TextMatrix(1, 2) = rsSerPat.Fields(2)
    fdata.TextMatrix(1, 3) = rsSerPat.Fields(3)
    'fdata.TextMatrix(1, 4) = rsSerPat.Fields(4)
    rsSerPat.MoveNext
    
    While Not rsSerPat.EOF
        fdata.AddItem rsSerPat.Fields(0) & vbTab & rsSerPat.Fields(1) & vbTab & rsSerPat.Fields(2) & vbTab & rsSerPat.Fields(3) & vbTab & rsSerPat.Fields(4)
        rsSerPat.MoveNext
    Wend
Else
MsgBox "Name/ID Doesn't Exist", vbCritical + vbOKOnly, "Invalid Name/ID"
fdata.clear
fdata.FixedRows = 1
fdata.Rows = 2
fdata.FormatString = "Service ID" & vbTab & "Service Name" & vbTab & "Duration" & vbTab & "Charge" & vbTab & "Notes"
    fdata.ColWidth(0) = 1500
    fdata.ColWidth(1) = 1800
    fdata.ColWidth(2) = 2000
    fdata.ColWidth(3) = 2000
    fdata.ColWidth(4) = 2000
End If
rsSerPat.Close
End If













End Sub

Private Sub cmbHospitalServiceID_Click()
cmbHospitalServiceID_Change
cmdSerSched_Click
End Sub

Private Sub cmbHospitalServiceID_KeyPress(KeyAscii As Integer)

flex = 20
If KeyAscii = 13 Then
fra.Visible = True
fdata.SetFocus
fdata.RowSel = 1
ElseIf KeyAscii = 27 Then
fra.Visible = False
End If


End Sub

Private Sub cmbPatientID_Change()
fdata.RowSel = 1


Dim rsSerPat As Recordset
Set rsSerPat = New ADODB.Recordset

flex = 10
fra.Visible = True
cmdPatSel.Visible = True
cmdDocIDSel.Visible = False



If ch = 0 Then

If Trim(cmbPatientID.Text) = "" Or Len(cmbPatientID.Text) = 0 Then
    rsSerPat.Open "select * from Patient_Details", cnPatients, adOpenDynamic, adLockPessimistic
Else
    rsSerPat.Open "select * from Patient_Details where Patient_ID like '" & UCase(Trim(cmbPatientID.Text)) & "%'", cnPatients, adOpenDynamic, adLockPessimistic
End If

If rsSerPat.RecordCount > 0 Then
    rsSerPat.MoveFirst
    fdata.clear
    fdata.FixedRows = 1
    fdata.Rows = 2
    fdata.FormatString = "Patient ID" & vbTab & "First Name" & vbTab & "Last Name" & vbTab & "Gender" & vbTab & "Address"
    fdata.ColWidth(0) = 1500
    fdata.ColWidth(1) = 1800
    fdata.ColWidth(2) = 1800
    fdata.ColWidth(3) = 1600
    fdata.ColWidth(4) = 1600
    fdata.TextMatrix(1, 0) = rsSerPat.Fields(0)
    fdata.TextMatrix(1, 1) = rsSerPat.Fields(1)
    fdata.TextMatrix(1, 2) = rsSerPat.Fields(2)
    fdata.TextMatrix(1, 3) = rsSerPat.Fields(3)
    fdata.TextMatrix(1, 4) = rsSerPat.Fields(4)
    rsSerPat.MoveNext
    
    While Not rsSerPat.EOF
        fdata.AddItem rsSerPat.Fields(0) & vbTab & rsSerPat.Fields(1) & vbTab & rsSerPat.Fields(2) & vbTab & rsSerPat.Fields(3) & vbTab & rsSerPat.Fields(4)
        rsSerPat.MoveNext
    Wend
Else
MsgBox "Name/ID Doesn't Exist", vbCritical + vbOKOnly, "Invalid Name/ID"
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
rsSerPat.Close
End If





End Sub

Private Sub cmbPatientID_Click()
cmbPatientID_Change
End Sub

Private Sub cmbPatientID_KeyPress(KeyAscii As Integer)
flex = 10
If KeyAscii = 13 Then
fra.Visible = True
fdata.SetFocus
fdata.RowSel = 1
ElseIf KeyAscii = 27 Then
fra.Visible = False

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




Set rsAddAppointment = New ADODB.Recordset
Set rsInfo = New ADODB.Recordset
Set rsSched = New ADODB.Recordset
Set rsAddBill = New ADODB.Recordset
Set rsPrevApp = New ADODB.Recordset


        ' Adding data to the Service_Appointment_Bill table
        strTime = DTPTime1.Value
        
                
        BID = Functions.UID(6, "OPBID_")    'Generate random Bill ID
        rsAddBill.Open "Select * from Service_Appointment_Bill", cnPatients, adOpenKeyset, adLockPessimistic
        While rsAddBill.EOF = False
            If rsAddBill(0) = BID Then  ' If Bill ID found Generate Another Bill ID
                BID = Functions.UID(6, "OPBID_")
                rsAddBill.MoveFirst
            End If
            rsAddBill.MoveNext
        Wend

            StrServiceID = cmbHospitalServiceID.Text
    
            rsAddAppointment.Open "Select * from Service_Appointment", cnPatients, adOpenKeyset, adLockPessimistic
            rsInfo.Open "Select * from Services where Channel_Service_ID='" & StrServiceID & "'", cnPatients, adOpenKeyset, adLockPessimistic
    
            If rsInfo.EOF = False Then
                strAmount = rsInfo![Charge_For_Service]
            End If

            
                    ' Assigning temporary Values
                    HospitalCharge = 200
                    Discount = 20
                    Debug.Print strAmount
                    GrandTotal = HospitalCharge + strAmount
                    NetValue = GrandTotal - Discount
                    AppBillID = BID
                    BillPatientID = cmbPatientID
        
                    ' End of BillData
            
            
            
    
            PID = Functions.UID(6, "SApp_")
            While rsAddAppointment.EOF = False
                If rsAddAppointment(0) = PID Then
                    PID = Functions.UID(6, "SApp_")
                    rsAddAppointment.MoveFirst
                    flag = True
                Else
                    flag = False
                End If
      
                rsAddAppointment.MoveNext
    
            Wend
        
            If cmbPatientID = "" Then
                MsgBox "Please enter a valid patient ID", vbCritical, "Out Patient Details"
                Exit Sub
            End If
            If cmbHospitalServiceID = "" Then
                MsgBox "Please enter a valid Service ID", vbCritical, "Out Patient Details"
                Exit Sub
            End If
            
            If DTPDate1.Value < Date Then
                MsgBox "Appointment Date Should Be Greater Than Current Date", vbCritical, "Invalid Date"
                Exit Sub
            End If
            If DTPDate1.Value = Date And strTime < Time Then
                MsgBox "Appointment Time Should Be Greater Than Current Time", vbCritical, "Invalid Date"
                rsAddBill.Close
                rsAddAppointment.Close
                rsInfo.Close
                Exit Sub
            End If
        
            Debug.Print StrServiceID
            rsSched.Open "Select * from Service_Schedule_Details where Service_ID='" & StrServiceID & "'", cnPatients, adOpenKeyset, adLockPessimistic
            
            ' Retreive Doctor Available days from the table
            While rsSched.EOF = False
                strAvailDays = strAvailDays & rsSched![Service_AvaiDate] & "..."
                strDocIn = strDocIn & rsSched![Service_Starts] & "..."
                strDocOut = strDocOut & rsSched![Service_Ends] & "..."
                rsSched.MoveNext
            Wend
            
            arrDays() = Split(strAvailDays, "...")
            strDIn() = Split(strDocIn, "...")
            strDOut() = Split(strDocOut, "...")
            
                   
            
            strDate = Left(Format(DTPDate1.Value, "dddd"), 3)
                       
            For i = 0 To UBound(arrDays)
                If InStr(1, arrDays(i), Left(Format(DTPDate1.Value, "dddd"), 3)) > 0 Then
                    AppPos = True ' Service is available on the selected date
                End If
            Next i
            
   

            fl = 1

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
                        Debug.Print "Service Starts        : " & strInTime
                        Debug.Print "Service Ends          : " & strOutTime
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
                Exit Sub
            End If

            
            rsPrevApp.Open "select * from Service_Appointment where Hospital_Service_ID='" & StrServiceID & "' and  Appointment_Date=#" & DTPDate1.Value & "# and Appointment_Time >= #" & strInTime & "# and  Appointment_time <= #" & strOutTime & "#", cnPatients, adOpenKeyset, adLockPessimistic
            NoOfSchedules = rsPrevApp.RecordCount
            
            While rsPrevApp.EOF = False
                Debug.Print "Current Appointments at :" & rsPrevApp![Appointment_Time]
                rsPrevApp.MoveNext
            Wend
            
            If NoOfSchedules > (newTime / 15) Then      '15 mins per patient
                MsgBox "No Space Available For This Appointment"
                Exit Sub
            End If
            
            
            
            ' Add data to the database if no errors occured
            
            If MsgBox("Are you sure you want to add this record to the database?", vbYesNo, "Add Service Appoinment") = vbYes Then
                
                cnPatients.Execute ("Insert into Service_Appointment values('" & PID & "','" & cmbPatientID & "','" & StrServiceID & "','" & DTPDate1.Value & "','" & strTime & "')")
                cnPatients.Execute ("Insert into Service_Appointment_Bill values('" & BID & "','" & PID & "','" & cmbPatientID & "','" & Format(Date, "mm/dd/yy") & "'," & strAmount & "," & HospitalCharge & "," & GrandTotal & "," & Discount & "," & Val(NetValue) & ")")
                
       
                
            rsAddBill.Close
            rsAddAppointment.Close
            rsInfo.Close
            rsPrevApp.Close
            rsSched.Close
                

                
                Unload Me
                frmSerAppoinmnetCharges.Show
                
            
            End If
        
        




End Sub

Private Sub cmdCancel_Click()
fra.Visible = False

End Sub

Private Sub cmdDocIDSel_Click()

Dim rsSerPat As Recordset
Set rsSerPat = New ADODB.Recordset

ch = 1

rsSerPat.Open "select * from Services where Channel_Service_ID like '" & UCase(Trim(cmbHospitalServiceID.Text)) & "%'", cnPatients, adOpenDynamic, adLockPessimistic
If rsSerPat.RecordCount > 0 Then
dupid = fdata.TextMatrix(fdata.Row, 0)
cmbHospitalServiceID.Text = dupid

Else
MsgBox "Select the Appropiate Data from the [[GridBox]]", vbInformation + vbOKOnly, "Error"

End If

rsSerPat.Close
dupid = 0
ch = 0
fra.Visible = False

End Sub



Private Sub cmdPatSel_Click()
Dim rsSerPat As Recordset
Set rsSerPat = New ADODB.Recordset

ch = 1

rsSerPat.Open "select * from Patient_Details where Patient_ID like '" & UCase(Trim(cmbPatientID.Text)) & "%'", cnPatients, adOpenDynamic, adLockPessimistic
If rsSerPat.RecordCount > 0 Then
dupid = fdata.TextMatrix(fdata.Row, 0)
cmbPatientID.Text = dupid
cmbHospitalServiceID.SetFocus

Else
MsgBox "Select the Appropiate Data from the [[GridBox]]", vbInformation + vbOKOnly, "Error"

End If

rsSerPat.Close
dupid = 0
ch = 0
fra.Visible = False
End Sub

Private Sub cmdSerSched_Click()



Dim rsSched As Recordset
Set rsSched = New ADODB.Recordset

   

rsSched.Open "Select * from Service_Schedule_Details where Service_ID='" & cmbHospitalServiceID & "'", cnPatients, adOpenKeyset, adLockPessimistic
            
            ' Retreive Doctor Available days from the table
If rsSched.EOF = False Then
        
    MSFlexGrid1.clear
    
    rsSched.MoveFirst
    
  
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
            
            
            
            
            rsSched.MoveNext
        Wend
    
    
        .TextMatrix(0, 0) = "Service ID"
        .TextMatrix(0, 1) = "Available Date"
        .TextMatrix(0, 2) = "Time In"
        .TextMatrix(0, 3) = "Time Out"
        .TextMatrix(0, 4) = "Schedule ID"
        .TextMatrix(0, 5) = "Notes"
   
        .FixedRows = 1
        .RowHeight(0) = .RowHeight(1) * 1.5
        
        .ColWidth(0) = 1500
        .ColWidth(1) = 2500
        .ColWidth(2) = 2500
        .ColWidth(3) = 2500
        .ColWidth(4) = 2500
        .ColWidth(5) = 2500
        
        'Functions.SizeColumns MSFlexGrid1, Me
        'Functions.SizeColumnHeaders MSFlexGrid1, Me
    End With
    
 Else
 Debug.Print "No Records Found"
 End If









End Sub

Private Sub Command1_Click()
frmService.Show
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub fdata_Click()
dupid = fdata.TextMatrix(fdata.Row, 0)
End Sub

Private Sub fdata_KeyPress(KeyAscii As Integer)

Dim rsSerPat As Recordset
Set rsSerPat = New ADODB.Recordset

If KeyAscii = 13 Then
    If flex = 20 Then
          dupid = fdata.TextMatrix(fdata.Row, 0)
        dupid1 = fdata.TextMatrix(fdata.Row, 0)
        customer_code = fdata.TextMatrix(fdata.Row, 0)
        ch = 1

        rsSerPat.Open "select * from Services where Channel_Service_ID like '" & cmbHospitalServiceID.Text & "%'", cnPatients, adOpenDynamic, adLockPessimistic
        If rsSerPat.RecordCount > 0 Then
            dupid = fdata.TextMatrix(fdata.Row, 0)
            cmbHospitalServiceID.Text = dupid
         
            
        Else
            MsgBox "Select the Appropiate Data from the [[GridBox]]", vbInformation + vbOKOnly, "Error"
        End If
            rsSerPat.Close
            dupid = 0
            ch = 0
            fra.Visible = False
    ElseIf flex = 10 Then
        dupid = fdata.TextMatrix(fdata.Row, 0)
        dupid1 = fdata.TextMatrix(fdata.Row, 0)
        ch = 1

        rsSerPat.Open "select * from Patient_Details where Patient_ID like '" & cmbPatientID.Text & "%'", cnPatients, adOpenDynamic, adLockPessimistic
        If rsSerPat.RecordCount > 0 Then
            dupid = fdata.TextMatrix(fdata.Row, 0)
            cmbPatientID.Text = dupid
            Debug.Print dupid
            cmbHospitalServiceID.SetFocus
            
        Else
            MsgBox "Select the Appropiate Data from the [[GridBox]]", vbInformation + vbOKOnly, "Error"
        End If
            rsSerPat.Close
            dupid = 0
            ch = 0
            fra.Visible = False
    End If
    
ElseIf KeyAscii = 27 Then
    If flex = 20 Then
        fra.Visible = False
        cmbHospitalServiceID.SetFocus
    Else
        fra.Visible = False
        cmbPatientID.SetFocus
    End If
End If









End Sub

Private Sub Form_Load()


    Me.WindowState = vbMaximized



Call Functions.DisableMenu

Dim SQL2 As String
Dim rsServices As Recordset
Dim rsPID As Recordset

Set rsServices = New ADODB.Recordset
Set rsPID = New ADODB.Recordset


If cnPatients.State = adStateOpen Then

SQL2 = "SELECT * FROM Services"

rsServices.Open SQL2, cnPatients, adOpenStatic, adLockPessimistic

    While rsServices.EOF = False
        cmbHospitalServiceID.AddItem rsServices(0)
        rsServices.MoveNext
    Wend
    
    rsServices.MoveFirst
    cmbHospitalServiceID.Text = rsServices(0)
    rsServices.Close
    
    rsPID.Open "select * from Patient_Details", cnPatients, adOpenStatic, adLockPessimistic
    
   
    
    While rsPID.EOF = False
        cmbPatientID.AddItem rsPID(0)
        rsPID.MoveNext
    Wend
   
   
   
    rsPID.MoveLast
    cmbPatientID.Text = rsPID(0)
    
    rsPID.Close
    
    Calendar1.Value = Date
    DTPDate1.Value = Date
    DTPTime1.Value = Time
    fra.Visible = False
    
       MSFlexGrid1.ColWidth(0) = 1500
    MSFlexGrid1.ColWidth(1) = 3500
    MSFlexGrid1.ColWidth(2) = 2000
    MSFlexGrid1.ColWidth(3) = 2000
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.ColWidth(5) = 2000
    
    
Else
    'when database connection error occurs
    MsgBox "Database Connection Error", vbCritical, "SD Hospitals PVT LTD"
End

End If

End Sub



Private Sub Form_Unload(Cancel As Integer)
Call Functions.EnableMenu
End Sub
