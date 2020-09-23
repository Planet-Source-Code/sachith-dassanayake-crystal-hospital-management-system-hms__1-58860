VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmAppoinments 
   BackColor       =   &H00FF8080&
   Caption         =   "Out Patient Appointments"
   ClientHeight    =   10740
   ClientLeft      =   1410
   ClientTop       =   2340
   ClientWidth     =   12525
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAppoinments.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10740
   ScaleWidth      =   12525
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid MSHFlexGrid1 
      Height          =   6975
      Left            =   480
      TabIndex        =   9
      Top             =   1560
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   12303
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      BackColorBkg    =   16744576
      HighLight       =   2
      AllowUserResizing=   3
      FormatString    =   "Appointment ID | Patien ID | Doctor / Service ID | Patient First Name | Patient Last Name | Contact Number | Date  | Time"
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "Appointments"
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
      Height          =   1215
      Left            =   480
      TabIndex        =   7
      Top             =   8760
      Width           =   11895
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Close"
         Height          =   855
         Left            =   5880
         Picture         =   "frmAppoinments.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add Appointment"
         Height          =   855
         Left            =   4320
         Picture         =   "frmAppoinments.frx":5CE6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "View Appointments"
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
      Height          =   3975
      Left            =   9000
      TabIndex        =   3
      Top             =   1440
      Width           =   3375
      Begin VB.OptionButton optApppointment 
         BackColor       =   &H00FF8080&
         Caption         =   "Service Appointments"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   6
         Top             =   3600
         Width           =   2295
      End
      Begin VB.OptionButton optApppointment 
         BackColor       =   &H00FF8080&
         Caption         =   "Doctor Appointments"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   5
         Top             =   3120
         Value           =   -1  'True
         Width           =   2415
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2310
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   14677179
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         MonthBackColor  =   -2147483648
         StartOfWeek     =   61538305
         CurrentDate     =   38328
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "View All Appointments"
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
      Height          =   2895
      Left            =   9000
      TabIndex        =   0
      Top             =   5640
      Width           =   3375
      Begin VB.CommandButton cmdServiceAppointment 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Service Appointments"
         Height          =   975
         Left            =   1200
         Picture         =   "frmAppoinments.frx":6181
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmdDoctorAppointment 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Doctor Appointments"
         Height          =   975
         Left            =   1200
         Picture         =   "frmAppoinments.frx":66F9
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   480
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   4
      ScreenHeight    =   1024
      ScreenWidth     =   1280
      ScreenHeightDT  =   1024
      ScreenWidthDT   =   1280
      FormHeightDT    =   11250
      FormWidthDT     =   12645
      FormScaleHeightDT=   10740
      FormScaleWidthDT=   12525
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "OUT PATIENT APPOINTMENTS "
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
      TabIndex        =   11
      Top             =   480
      Width           =   6195
   End
End
Attribute VB_Name = "frmAppoinments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private AppointmentSelect As Integer
Private rsAppointment As ADODB.Recordset 'Appointment table

Private Sub cmdDoctorAppointment_Click()
Dim SQL As String
Dim icol As Integer
Set rsAppointment = New ADODB.Recordset
'create sql statement

SQL = "SELECT * FROM Doctor_Appointment ORDER BY Appointment_Date, Appointment_Time "

Set rsAppointment.ActiveConnection = cnPatients
rsAppointment.Open SQL

With MSHFlexGrid1
    .clear
    .Rows = 1
    .Cols = rsAppointment.Fields.Count
  


    While Not rsAppointment.EOF
        .Rows = .Rows + 1
        .Row = .Rows - 1


        For icol = 0 To rsAppointment.Fields.Count - 1
            .Col = icol
            .Text = rsAppointment(icol) & ""
        Next
        rsAppointment.MoveNext
    Wend

    .TextMatrix(0, 0) = "Appointment ID"
    .TextMatrix(0, 1) = "Patient ID"
    .TextMatrix(0, 2) = "Doctor ID"
    .TextMatrix(0, 3) = "Appointment Date"
    .TextMatrix(0, 4) = "Appointment Time"

    .FixedRows = 1
    .RowHeight(0) = .RowHeight(1) * 1.5
    
    '.FontWidthFixed = 5
 
    
End With

Functions.SizeColumns MSHFlexGrid1, Me

rsAppointment.Close

End Sub


Private Sub cmdServiceAppointment_Click()
Dim SQL As String
Dim icol As Integer
Set rsAppointment = New ADODB.Recordset
'create sql statement

SQL = "SELECT * FROM Service_Appointment ORDER BY Appointment_Date, Appointment_Time "

Set rsAppointment.ActiveConnection = cnPatients
rsAppointment.Open SQL

With MSHFlexGrid1
    .clear
    .Rows = 1
    .Cols = rsAppointment.Fields.Count
  

    While Not rsAppointment.EOF
        .Rows = .Rows + 1
        .Row = .Rows - 1


        For icol = 0 To rsAppointment.Fields.Count - 1
            .Col = icol
            .Text = rsAppointment(icol) & ""
        Next
        rsAppointment.MoveNext
    Wend

    .TextMatrix(0, 0) = "Appointment Number"
    .TextMatrix(0, 1) = "Patient ID"
    .TextMatrix(0, 2) = "Service ID"
    .TextMatrix(0, 3) = "Appointment Date"
    .TextMatrix(0, 4) = "Appointment Time"

    .FixedRows = 1
    .RowHeight(0) = .RowHeight(1) * 1.5
 
    
End With

Functions.SizeColumns MSHFlexGrid1, Me

rsAppointment.Close

End Sub


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call Functions.DisableMenu
Call FillTable
MonthView1.Value = Now
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Functions.EnableMenu
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
Dim SQL As String
Dim RowNo As Integer
Set rsAppointment = New ADODB.Recordset

MSHFlexGrid1.clear

    If AppointmentSelect = 0 Then
        'create sql statement
        SQL = "SELECT * From Doctor_Appointment, Patient_Details WHERE Doctor_Appointment.Appointment_Date =#" & SQLDate(DateClicked) & "#" & " and Doctor_Appointment.Patient_ID=Patient_Details.Patient_ID ORDER BY Appointment_Date, Appointment_Time" & ";"
    ElseIf AppointmentSelect = 1 Then
        'create sql statement
        SQL = "SELECT * From Service_Appointment, Patient_Details WHERE Appointment_Date =#" & SQLDate(DateClicked) & "#" & " and Service_Appointment.Patient_ID=Patient_Details.Patient_ID ORDER BY Appointment_Date, Appointment_Time" & ";"
    End If



Set rsAppointment.ActiveConnection = cnPatients

rsAppointment.Open SQL

If Not rsAppointment.EOF Then ' If results found
    With MSHFlexGrid1
        .clear
        .Rows = 1
        .Cols = rsAppointment.Fields.Count
        
        RowNo = 1

            While Not rsAppointment.EOF
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .TextMatrix(RowNo, 0) = rsAppointment![Appointment_ID]
                .TextMatrix(RowNo, 1) = rsAppointment![First_Name]
                .TextMatrix(RowNo, 2) = rsAppointment![Last_Name]
                .TextMatrix(RowNo, 3) = rsAppointment![Gender]
                .TextMatrix(RowNo, 4) = rsAppointment![address]
                .TextMatrix(RowNo, 5) = rsAppointment![Telephone]
                
                If AppointmentSelect = 0 Then
                    .TextMatrix(RowNo, 6) = rsAppointment![Doctor_ID]
                ElseIf AppointmentSelect = 1 Then
                    .TextMatrix(RowNo, 6) = rsAppointment![Hospital_SErvice_ID]
                End If
                
                .TextMatrix(RowNo, 7) = rsAppointment![Appointment_Time]
            
                rsAppointment.MoveNext
                RowNo = RowNo + 1
            Wend

        ' Set Column Headers

        .TextMatrix(0, 0) = "Appointment ID"
        .TextMatrix(0, 1) = "First Name"
        .TextMatrix(0, 2) = "Last Name"
        .TextMatrix(0, 3) = "Gender"
        .TextMatrix(0, 4) = "Address"
        .TextMatrix(0, 5) = "Phone Number"
        .TextMatrix(0, 6) = "Doctor / Service"
        .TextMatrix(0, 7) = "Time"
        .FixedRows = 1
        .RowHeight(0) = .RowHeight(1) * 1.5
        
        Functions.SizeColumns MSHFlexGrid1, Me

    End With

Else ' If no Results Found

   With MSHFlexGrid1
        
        .Cols = 8
        .clear
        
        .TextMatrix(0, 0) = "Appointment ID"
        .TextMatrix(0, 1) = "First Name"
        .TextMatrix(0, 2) = "Last Name"
        .TextMatrix(0, 3) = "Gender"
        .TextMatrix(0, 4) = "Address"
        .TextMatrix(0, 5) = "Phone Number"
        .TextMatrix(0, 6) = "Doctor / Service"
        .TextMatrix(0, 7) = "Time"
        
        .FixedRows = 1
        .RowHeight(0) = .RowHeight(1) * 1.5

        '.FontWidth = 6
        Functions.SizeColumnHeaders MSHFlexGrid1, Me
    End With



End If

rsAppointment.Close

End Sub



Private Sub optApppointment_Click(Index As Integer)
    
    Select Case (Index)
        Case "0" ' Doctor Appointments
           AppointmentSelect = 0
           Call MonthView1_DateClick(MonthView1.Value)
           
           
        Case "1" 'Other Hospital Services Appointments
            AppointmentSelect = 1
            Call MonthView1_DateClick(MonthView1.Value)
           

        Case Else 'None
            
    End Select

End Sub



Public Sub FillTable()

Dim rsAppointment2 As Recordset
Dim SQL1 As String
Dim SQL2 As String
Dim icol As Integer

Set rsAppointment = New ADODB.Recordset
Set rsAppointment2 = New ADODB.Recordset

'create sql statement

SQL1 = "SELECT * FROM Doctor_Appointment WHERE Doctor_Appointment.Appointment_Date =#" & Format(Date, "dd-MMM-yyyy") & "# "
SQL2 = "SELECT * FROM Service_Appointment WHERE Service_Appointment.Appointment_Date =#" & Format(Date, "dd-MMM-yyyy") & "#"
Set rsAppointment.ActiveConnection = cnPatients
Set rsAppointment2.ActiveConnection = cnPatients

rsAppointment.Open SQL1
rsAppointment2.Open SQL2



If rsAppointment.EOF = False Or rsAppointment2.EOF = False Then

With MSHFlexGrid1
    .clear
    .Rows = 1
    .Cols = rsAppointment.Fields.Count
   
  


    While Not rsAppointment.EOF
        .Rows = .Rows + 1
        .Row = .Rows - 1


        For icol = 0 To rsAppointment.Fields.Count - 1
            .Col = icol
            .Text = rsAppointment(icol) & ""
        Next
        rsAppointment.MoveNext
    Wend
    
    
    While Not rsAppointment2.EOF
        .Rows = .Rows + 1
        .Row = .Rows - 1


        For icol = 0 To rsAppointment2.Fields.Count - 1
            .Col = icol
            .Text = rsAppointment2(icol) & ""
        Next
        rsAppointment2.MoveNext
    Wend

    .TextMatrix(0, 0) = "Appointment ID"
    .TextMatrix(0, 1) = "Patient ID"
    .TextMatrix(0, 2) = "Doctor / Service ID"
    .TextMatrix(0, 3) = "Appointment Date"
    .TextMatrix(0, 4) = "Appointment Time"

    .FixedRows = 1
    .RowHeight(0) = .RowHeight(1) * 1.5

End With

Functions.SizeColumns MSHFlexGrid1, Me

rsAppointment.Close
End If


End Sub
