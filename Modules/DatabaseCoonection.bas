Attribute VB_Name = "DatabaseConnection"
Option Explicit
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Public cnPatients As ADODB.Connection  'cnPatient is database name

Public AppBillID As String
Public strAmount As Currency
Public AppointmentCharge As Currency
Public HospitalCharge As Currency
Public GrandTotal As Currency
Public Discount As Currency
Public NetValue As Currency
Public BillPatientID As String

Public UserCategory As Integer
Public User As String
Public LogDate As Date
Public LogTime As Date
Public AppState As Integer
Public appRegistered As Boolean



Sub Main()
    On Error Resume Next
    
   
     
     Dim crackkey As String
    'Read Registry for previous settings stored
    crackkey = GetSetting(App.Title, "Settings", "CHECK")
    If crackkey = "" Then
        MsgBox "You are not using License Version of Crystal Hospital Management System" & vbCrLf & "Please Register The Application ", vbInformation, "Authentication Check"
        appRegistered = False
    Else
        appRegistered = True
        
    End If
    
    If App.PrevInstance = True Then
        MsgBox "Crystal Hospital Management System is already open", vbInformation, "Crystal Hospital Management System"
        Exit Sub
    End If
    
    
    AppState = 1
    Dim sConnect As String
    
    Set cnPatients = New ADODB.Connection
    
    sConnect = "provider=MSDataShape;Data provider=Microsoft.Jet.OLEDB.4.0;data source=" & App.Path & "\database\HMS.mdb;"
    
    cnPatients.CursorLocation = adUseClient
    cnPatients.Open (sConnect)
    
    If Not cnPatients.State = adStateOpen Then
        MsgBox "Database Error. Please Check the database and try again", vbCritical, "Database Error"
        End
    End If
  
   frmSplashScreen.Show
    'MDIMain.Show
    'frmDoctorSchedule.Show
    'frmServiceSchedule.Show
     
     
End Sub
Public Sub disMenu()
MDIMain.EmployeeMng.Enabled = False
    If appRegistered = True Then
        MDIMain.register1.Enabled = False
    End If
End Sub


Public Sub Privilages()
On Error Resume Next

Dim pctl As Control

Select Case UserCategory

' If the user is Guest
Case 0

With MDIMain
    .PharmacyMng.Visible = False
    .patients.Visible = False
    .manegement.Visible = False
   
    .AddDoc.Visible = False
    .AddSer.Visible = False
    .Rep.Visible = False
    .tools.Visible = False
      
    .Settings.Visible = False
    .CoolBar1.Visible = False
    .backup.Visible = False
    
End With



' If the user is Administrator
Case 1

For Each pctl In MDIMain.Controls
    pctl.Visible = True
Next
MDIMain.CoolBar1.Bands(1).Visible = True
MDIMain.CoolBar1.Bands(2).Visible = True
  

' If the user is Employee Manager
Case 2
With MDIMain
For Each pctl In MDIMain.Controls
    pctl.Visible = True
Next

.PharmacyMng.Visible = False
.patients.Visible = False
.HosManage.Visible = False
.EmployeeMng.Visible = True

.AddDoc.Visible = False
.AddSer.Visible = False
.rptHospital.Visible = False
.rptPharmacy.Visible = False
.rptEmployee.Visible = True
.PatientReports.Visible = False
.CoolBar1.Bands(2).Visible = False
.CoolBar1.Bands(1).Visible = False
.AddnewUser.Visible = False
    

End With


' If the user is Patient Manager
Case 3
For Each pctl In MDIMain.Controls
    pctl.Visible = True
Next

With MDIMain
.PharmacyMng.Visible = False
.patients.Visible = True
.HosManage.Visible = False
.EmployeeMng.Visible = False

.rptHospital.Visible = True
.rptPharmacy.Visible = False
.rptEmployee.Visible = False


.CoolBar1.Bands(2).Visible = False
.CoolBar1.Bands(1).Visible = True
.AddnewUser.Visible = False
.backup.Visible = False
End With



' If the user is Pharmacy Manager
Case 4
For Each pctl In MDIMain.Controls
    pctl.Visible = True
Next

With MDIMain
.PharmacyMng.Visible = True
.patients.Visible = False
.HosManage.Visible = False
.EmployeeMng.Visible = False
.AddDoc.Visible = False
.AddSer.Visible = False

.rptHospital.Visible = False
.rptPharmacy.Visible = True
.rptEmployee.Visible = False
.PatientReports.Visible = False

.CoolBar1.Bands(2).Visible = True
.CoolBar1.Bands(1).Visible = False
.AddnewUser.Visible = False
.backup.Visible = False
   
End With

' If the user is Manager
Case 5
For Each pctl In MDIMain.Controls
    pctl.Visible = True
Next

With MDIMain
.PharmacyMng.Visible = False
.patients.Visible = False
.HosManage.Visible = True
.EmployeeMng.Visible = False

.rptHospital.Visible = True
.rptPharmacy.Visible = True
.rptEmployee.Visible = True
.PatientReports.Visible = False

.CoolBar1.Bands(2).Visible = False
.CoolBar1.Bands(1).Visible = False
.AddnewUser.Visible = False
.backup.Visible = False
End With

' If the user is Employee User
Case 6
For Each pctl In MDIMain.Controls
    pctl.Visible = True
Next

With MDIMain
    .EmployeeMng.Visible = True
    
    .PharmacyMng.Visible = False
    .patients.Visible = False
    .HosManage.Visible = False
    .Rep.Visible = False
    .AddDoc.Visible = False
    .AddSer.Visible = False
    .CoolBar1.Bands(2).Visible = False
    .CoolBar1.Bands(1).Visible = False
    .AddnewUser.Visible = False
    .PatientReports.Visible = False
    .backup.Visible = False
        
End With


' If the user is Patient User
Case 7
For Each pctl In MDIMain.Controls
    pctl.Visible = True
Next

With MDIMain
    .patients.Visible = True
    .PharmacyMng.Visible = False
    .HosManage.Visible = False
    .EmployeeMng.Visible = False
    .manegement.Visible = False
    .Rep.Visible = False
   
    .AddDoc.Visible = False
    .AddSer.Visible = False
    
    .CoolBar1.Bands(2).Visible = False
    .CoolBar1.Bands(1).Visible = True
    .AddnewUser.Visible = False
    .PatientReports.Visible = False
    .backup.Visible = False
      
End With



' If the user is Pharmacy User
Case 8

For Each pctl In MDIMain.Controls
    pctl.Visible = True
Next

With MDIMain
    .PharmacyMng.Visible = True
    .patients.Visible = False
    .HosManage.Visible = False
    .EmployeeMng.Visible = False
    .Rep.Visible = False
    .AddDoc.Visible = False
    .AddSer.Visible = False
    .Purchases.Visible = False
    .CoolBar1.Bands(2).Visible = True
    .CoolBar1.Bands(1).Visible = False
    .AddnewUser.Visible = False
    .PatientReports.Visible = False
    .backup.Visible = False
       
End With




End Select


End Sub
