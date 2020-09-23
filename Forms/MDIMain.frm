VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm MDIMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0FF&
   Caption         =   " Crystal Hospital Management System"
   ClientHeight    =   11580
   ClientLeft      =   2610
   ClientTop       =   1605
   ClientWidth     =   13920
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "MDIMain.frx":068A
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3000
      Top             =   1200
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3480
      Top             =   1200
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   11205
      Width           =   13920
      _ExtentX        =   24553
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "1/30/2005"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Text            =   "CAPS"
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Text            =   "NUM"
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Text            =   "INS"
            TextSave        =   "INS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13920
      _ExtentX        =   24553
      _ExtentY        =   767
      BandCount       =   2
      _CBWidth        =   13920
      _CBHeight       =   435
      _Version        =   "6.0.8169"
      Caption1        =   "Patients"
      Child1          =   "Picture1"
      MinHeight1      =   375
      Width1          =   9495
      NewRow1         =   0   'False
      Caption2        =   "Pharmacy"
      Child2          =   "Picture2"
      MinHeight2      =   375
      Width2          =   5175
      NewRow2         =   0   'False
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         FillColor       =   &H80000004&
         ForeColor       =   &H80000004&
         Height          =   375
         Left            =   10590
         ScaleHeight     =   375
         ScaleWidth      =   3240
         TabIndex        =   4
         Top             =   30
         Width           =   3240
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   330
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   582
            ButtonWidth     =   2381
            ButtonHeight    =   582
            Wrappable       =   0   'False
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "imlMain"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   3
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Customers"
                  Description     =   "Add New Out Patient"
                  Object.ToolTipText     =   "Add New Customer"
                  ImageIndex      =   10
                  BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                     NumButtonMenus  =   1
                     BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     EndProperty
                  EndProperty
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Sales Invoice"
                  Description     =   "Add New Doctor Appointment"
                  Object.ToolTipText     =   "Add New Sale"
                  ImageIndex      =   11
               EndProperty
            EndProperty
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         FillColor       =   &H80000004&
         ForeColor       =   &H80000004&
         Height          =   375
         Left            =   945
         ScaleHeight     =   375
         ScaleWidth      =   8520
         TabIndex        =   2
         Top             =   30
         Width           =   8520
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   330
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   8520
            _ExtentX        =   15028
            _ExtentY        =   582
            ButtonWidth     =   2461
            ButtonHeight    =   582
            Wrappable       =   0   'False
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "imlMain"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   7
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Out Patient"
                  Description     =   "Add New Out Patient"
                  Object.ToolTipText     =   "Add New Out Patient"
                  ImageIndex      =   6
                  BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                     NumButtonMenus  =   1
                     BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     EndProperty
                  EndProperty
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Appointments"
                  Description     =   "Add New Doctor Appointment"
                  Object.ToolTipText     =   "Add New Appointments"
                  ImageIndex      =   4
                  Style           =   5
                  BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                     NumButtonMenus  =   2
                     BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Text            =   "Doctor Appointment"
                     EndProperty
                     BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Text            =   "Medical Appointment"
                     EndProperty
                  EndProperty
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "In Patient"
                  Description     =   "Add In Patient Details"
                  Object.ToolTipText     =   "Add New In Patient"
                  ImageIndex      =   7
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Admission"
                  Description     =   "Add Admission Details"
                  Object.ToolTipText     =   "New Admission"
                  ImageIndex      =   8
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Payments"
                  Object.ToolTipText     =   "Payments and Billing"
                  ImageIndex      =   9
                  Style           =   5
                  BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                     NumButtonMenus  =   4
                     BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Text            =   "In Patient Bill Payments"
                     EndProperty
                     BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Text            =   "-"
                     EndProperty
                     BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Text            =   "Doctor Appointment Payments"
                     EndProperty
                     BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                        Text            =   "Service Appointment Payments"
                     EndProperty
                  EndProperty
               EndProperty
            EndProperty
         End
      End
   End
   Begin MSComctlLib.ImageList imlMain 
      Left            =   4560
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":7E53
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":802D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":8665
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":897F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":9E81
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":A05B
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":A52E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":A9F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":AE9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":B368
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMain.frx":B82B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Main 
      Caption         =   "&Main"
      Begin VB.Menu Doctor 
         Caption         =   "Doctor Details"
         Begin VB.Menu AddDoc 
            Caption         =   "Add Doctors"
            Shortcut        =   ^D
         End
         Begin VB.Menu ViewDoc 
            Caption         =   "Dispaly Doctors"
            Shortcut        =   ^E
         End
         Begin VB.Menu sep 
            Caption         =   "-"
         End
         Begin VB.Menu SearchDoc 
            Caption         =   "Search Doctors"
         End
      End
      Begin VB.Menu Service 
         Caption         =   "Hospital Services"
         Begin VB.Menu AddSer 
            Caption         =   "Add Hospital Service"
            Shortcut        =   ^S
         End
         Begin VB.Menu DisSer 
            Caption         =   "Display Hospital Services"
            Shortcut        =   ^T
         End
         Begin VB.Menu sep2 
            Caption         =   "-"
         End
         Begin VB.Menu SearchSer 
            Caption         =   "Search Hospital Services"
         End
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu backup 
         Caption         =   "Backup Database"
         Shortcut        =   ^B
      End
      Begin VB.Menu sep34 
         Caption         =   "-"
      End
      Begin VB.Menu logoff 
         Caption         =   "Log Off"
         Shortcut        =   ^{F11}
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu patients 
      Caption         =   "&Patients"
      Begin VB.Menu OPatients 
         Caption         =   "&Out Patients"
         Begin VB.Menu OPDetails 
            Caption         =   "Out Patient Details"
            Begin VB.Menu addOPatient 
               Caption         =   "Add New Patient"
               Shortcut        =   ^O
            End
            Begin VB.Menu ViewOPatient 
               Caption         =   "View All"
               Shortcut        =   ^P
            End
         End
         Begin VB.Menu sep29 
            Caption         =   "-"
         End
         Begin VB.Menu OutAppointment 
            Caption         =   "Docotor Appointment"
            Begin VB.Menu DocApp 
               Caption         =   "Add Doctor Appointments"
               Shortcut        =   ^Q
            End
            Begin VB.Menu DelDocApp 
               Caption         =   "Cancel Doctor Appointment"
            End
            Begin VB.Menu sep22 
               Caption         =   "-"
            End
            Begin VB.Menu SearchDocApp 
               Caption         =   "Search Doctor Appointment"
            End
         End
         Begin VB.Menu serApp 
            Caption         =   "Service Appointments"
            Begin VB.Menu AddSerApp 
               Caption         =   "Add Service Appointment"
               Shortcut        =   ^R
            End
            Begin VB.Menu Del 
               Caption         =   "Cancel Service Appointment"
            End
            Begin VB.Menu sep21 
               Caption         =   "-"
            End
            Begin VB.Menu SearchSerApp 
               Caption         =   "Search Service Appointment"
            End
         End
         Begin VB.Menu sep31 
            Caption         =   "-"
         End
         Begin VB.Menu opPrescription 
            Caption         =   "Patient Prescription"
            Begin VB.Menu addnewPrescription 
               Caption         =   "Add Prescription"
               Shortcut        =   ^F
            End
            Begin VB.Menu viewOPHistory 
               Caption         =   "View Patient History"
            End
         End
         Begin VB.Menu sep30 
            Caption         =   "-"
         End
         Begin VB.Menu OPBillPay 
            Caption         =   "Bill Payments"
            Begin VB.Menu ViewOBill 
               Caption         =   "View Doctor Appointment Patient Bill"
               Shortcut        =   ^G
            End
            Begin VB.Menu sep7 
               Caption         =   "-"
            End
            Begin VB.Menu AddOPay 
               Caption         =   "View Service Appointment Patient Bill"
               Shortcut        =   ^H
            End
         End
      End
      Begin VB.Menu sep25 
         Caption         =   "-"
      End
      Begin VB.Menu IPateint 
         Caption         =   "&In Patients"
         Begin VB.Menu InAddmision 
            Caption         =   "Admision"
            Begin VB.Menu AddIPatient 
               Caption         =   "Add Patient Details"
               Shortcut        =   ^I
            End
            Begin VB.Menu AddOGuard 
               Caption         =   "Add Guardian Details"
               Shortcut        =   ^J
            End
            Begin VB.Menu RegIPatient 
               Caption         =   "Registration"
               Shortcut        =   ^K
            End
            Begin VB.Menu sep9 
               Caption         =   "-"
            End
            Begin VB.Menu viewAddmission 
               Caption         =   "View Admission List"
            End
         End
         Begin VB.Menu ITreat 
            Caption         =   "Treatments"
            Begin VB.Menu AddDocVisits 
               Caption         =   "Add Doctor Visits"
            End
            Begin VB.Menu ViewDocVisits 
               Caption         =   "View Doctor Visits"
            End
            Begin VB.Menu sep16 
               Caption         =   "-"
            End
            Begin VB.Menu addmediciine 
               Caption         =   "Add Medicinal Details"
            End
            Begin VB.Menu ViewMedDetails 
               Caption         =   "View Medicinal Details"
            End
            Begin VB.Menu sep15 
               Caption         =   "-"
            End
            Begin VB.Menu AddIPService 
               Caption         =   "Add Service Details"
            End
            Begin VB.Menu ViewSerDetails 
               Caption         =   "View Service Details"
            End
         End
         Begin VB.Menu InDischarge 
            Caption         =   "Discharge"
            Begin VB.Menu IPDischarge 
               Caption         =   "Discharge Patient"
            End
            Begin VB.Menu sep10 
               Caption         =   "-"
            End
            Begin VB.Menu ViewDischarge 
               Caption         =   "View Discharge List"
            End
         End
         Begin VB.Menu billpay 
            Caption         =   "Bill Payments"
            Begin VB.Menu IViewBill 
               Caption         =   "View Patient Bill"
            End
            Begin VB.Menu editPatient 
               Caption         =   "Edit Patient Bill"
            End
            Begin VB.Menu sep4 
               Caption         =   "-"
            End
            Begin VB.Menu IAddNewPay 
               Caption         =   "Add New Payment"
            End
            Begin VB.Menu sep3 
               Caption         =   "-"
            End
            Begin VB.Menu ISerPay 
               Caption         =   "Search Payment"
            End
         End
      End
   End
   Begin VB.Menu manegement 
      Caption         =   "Management"
      Begin VB.Menu HosManage 
         Caption         =   "&Hopital Management"
         Begin VB.Menu DocMng 
            Caption         =   "Doctor Management"
            Begin VB.Menu docSched 
               Caption         =   "Doctor Schedule"
            End
         End
         Begin VB.Menu ServiceMng 
            Caption         =   "Hospital Service Management"
            Begin VB.Menu sersched 
               Caption         =   "Service Scehdule"
            End
         End
         Begin VB.Menu sep14 
            Caption         =   "-"
         End
         Begin VB.Menu RoomMng 
            Caption         =   "Room Management"
            Begin VB.Menu addRoom 
               Caption         =   "Add New Room"
            End
            Begin VB.Menu DisplayRoom 
               Caption         =   "View Rooms"
            End
            Begin VB.Menu addRoomType 
               Caption         =   "Add Room Type"
            End
         End
         Begin VB.Menu WardMng 
            Caption         =   "Ward Management"
            Begin VB.Menu addWard 
               Caption         =   "Add New Ward"
            End
            Begin VB.Menu DisplayWard 
               Caption         =   "View Ward Details"
            End
         End
         Begin VB.Menu BedMng 
            Caption         =   "Bed Management"
            Begin VB.Menu addBed 
               Caption         =   "Add Bed Details"
            End
            Begin VB.Menu BedAvail 
               Caption         =   "Bed Avaiabilty"
            End
         End
      End
      Begin VB.Menu PharmacyMng 
         Caption         =   "Pharmacy Management"
         Begin VB.Menu sale 
            Caption         =   "Sale"
            Begin VB.Menu Invoice 
               Caption         =   "Invoice"
            End
            Begin VB.Menu Sep19 
               Caption         =   "-"
            End
            Begin VB.Menu Customers 
               Caption         =   "Customers"
            End
         End
         Begin VB.Menu Purchases 
            Caption         =   "Purchases"
            Begin VB.Menu PInvoice 
               Caption         =   "Purchase Invoice"
            End
            Begin VB.Menu Sep20 
               Caption         =   "-"
            End
            Begin VB.Menu Products 
               Caption         =   "Products"
            End
            Begin VB.Menu Categories 
               Caption         =   "Categories"
            End
            Begin VB.Menu Suppliers 
               Caption         =   "Suppliers"
            End
         End
         Begin VB.Menu sep27 
            Caption         =   "-"
         End
         Begin VB.Menu showNavigation 
            Caption         =   "Show Menu"
         End
      End
      Begin VB.Menu EmployeeMng 
         Caption         =   "Employee Management"
         Begin VB.Menu AddempDetails 
            Caption         =   "Add Employee Details"
         End
         Begin VB.Menu EmpSal 
            Caption         =   "Employee Salary Calculation"
         End
         Begin VB.Menu DocSalcal 
            Caption         =   "Doctor Salary Calculation"
         End
         Begin VB.Menu addDepartment 
            Caption         =   "Add Departments"
         End
         Begin VB.Menu sep23 
            Caption         =   "-"
         End
         Begin VB.Menu showempMenu 
            Caption         =   "Show Menu"
         End
      End
   End
   Begin VB.Menu Rep 
      Caption         =   "&Reports"
      Begin VB.Menu rptMngs 
         Caption         =   "Managerial Reports"
         Begin VB.Menu PatientReports 
            Caption         =   "Patient Report"
            Begin VB.Menu IPMEdIssue 
               Caption         =   "In Patient Medicine Issue"
            End
            Begin VB.Menu IPMedServices 
               Caption         =   "In Patient Medical Services"
            End
            Begin VB.Menu IPDocvisits 
               Caption         =   "In Patient Doctor Visits"
            End
         End
         Begin VB.Menu rptHospital 
            Caption         =   "Hospital Reports"
            Begin VB.Menu RoomRpt 
               Caption         =   "Room Report"
            End
            Begin VB.Menu wardrpt 
               Caption         =   "Ward Report"
            End
            Begin VB.Menu BedRpt 
               Caption         =   "Bed Reports"
            End
            Begin VB.Menu sep28 
               Caption         =   "-"
            End
            Begin VB.Menu MedSerShedrpt 
               Caption         =   "Medical Services Schedule"
            End
            Begin VB.Menu docSchedRpt 
               Caption         =   "Doctor Shedule"
            End
         End
         Begin VB.Menu rptPharmacy 
            Caption         =   "Pharmacy Reports"
            Begin VB.Menu rptSales 
               Caption         =   "Sales"
            End
            Begin VB.Menu rptPurchase 
               Caption         =   "Purchases"
            End
            Begin VB.Menu rptCustomers 
               Caption         =   "Customers"
            End
         End
         Begin VB.Menu rptEmployee 
            Caption         =   "Employee Reports"
         End
      End
      Begin VB.Menu rptLogs 
         Caption         =   "Log Reports"
      End
   End
   Begin VB.Menu tools 
      Caption         =   "&Tools"
      Begin VB.Menu chngPass 
         Caption         =   "Change Password"
      End
      Begin VB.Menu AddnewUser 
         Caption         =   "Add New User"
      End
   End
   Begin VB.Menu Settings 
      Caption         =   "&Settings"
      Begin VB.Menu hosdetails 
         Caption         =   "Hospital Details"
      End
      Begin VB.Menu sidebar 
         Caption         =   "Side Bar"
         Checked         =   -1  'True
      End
      Begin VB.Menu appInterface 
         Caption         =   "Application Preferences"
      End
      Begin VB.Menu sep18 
         Caption         =   "-"
      End
      Begin VB.Menu appSkin 
         Caption         =   "Application Skin"
         Begin VB.Menu defualt 
            Caption         =   "Default"
         End
         Begin VB.Menu macgrey 
            Caption         =   "Mac Grey"
         End
         Begin VB.Menu xpblue 
            Caption         =   "XP Blue"
         End
         Begin VB.Menu coolgreen 
            Caption         =   "Cool Green"
         End
         Begin VB.Menu LightB 
            Caption         =   "Light Brown"
         End
         Begin VB.Menu lightV 
            Caption         =   "Light Violet"
         End
         Begin VB.Menu winclassic 
            Caption         =   "Win Classic"
         End
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu content 
         Caption         =   "Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu sep32 
         Caption         =   "-"
      End
      Begin VB.Menu shortCuts 
         Caption         =   "Short Cut Keys"
         Shortcut        =   {F2}
      End
      Begin VB.Menu credits 
         Caption         =   "Credits"
         Shortcut        =   +^{F9}
      End
      Begin VB.Menu sep33 
         Caption         =   "-"
      End
      Begin VB.Menu register1 
         Caption         =   "Register Crystal HMS"
         Shortcut        =   +^{F11}
      End
      Begin VB.Menu sep13 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "About Crystal HMS"
         Shortcut        =   +^{F12}
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub about_Click()
frmAbout.Show
End Sub

Private Sub addBed_Click()
frmAddBedDetails.Show
End Sub

Private Sub addDepartment_Click()
frm_settngs.Show
End Sub

Private Sub addDoc_Click()
frmAddDoctorDetails.Show
End Sub

Private Sub AddDocVisits_Click()
frmDoctorVisit.Show
End Sub


Private Sub AddempDetails_Click()
frm_add_employee.Show
End Sub

Private Sub AddIPatient_Click()
frmInPatientDetails.Show
End Sub

Private Sub AddIPService_Click()
frmInPatientServices.Show
End Sub

Private Sub addmediciine_Click()
frmInPatientMedicine.Show
End Sub

Private Sub addnewPrescription_Click()
frmOutPatientTreatments.Show

End Sub

Private Sub AddnewUser_Click()
frmUserDetails.Show
End Sub

Private Sub AddOGuard_Click()
frmGuardianDetails.Show
End Sub

Private Sub addOPatient_Click()
frmAddOutPatientDetails.Show
End Sub

Private Sub AddOPay_Click()
frmOPSerBillPayments.Show
End Sub

Private Sub addRoom_Click()
frmAddRoom.Show
End Sub

Private Sub addRoomType_Click()
frmAddRoomType.Show
End Sub

Private Sub AddSer_Click()
frmAddServices.Show
End Sub

Private Sub AddSerApp_Click()
frmAddSerAppointments.Show
End Sub


Private Sub addWard_Click()
frmAddWard.Show
End Sub

Private Sub appInterface_Click()
frmAppPreferences.Show
End Sub

Private Sub backup_Click()
frmBackUp.Show
End Sub

Private Sub BedAvail_Click()
frmBEDDisplay.Show
End Sub

Private Sub BedRpt_Click()
frmbedwardreport.Show
End Sub

Private Sub Categories_Click()
FrmCategories.Show
End Sub

Public Sub chngPass_Click()
frmChangePassword.Show
End Sub

Private Sub content_Click()
  'View help files
   Dim strHFile As String
   
   strHFile = App.Path & "\Help\CHMS.chm"
   
   ShellExecute Me.hWnd, "open", strHFile, "", "", vbNormalFocus
End Sub

Private Sub coolgreen_Click()
Call select_color_type(3)
sys_color = "3"
End Sub

Private Sub credits_Click()
frmCredits.Show
End Sub

Private Sub Customers_Click()
FrmCustomers.Show
End Sub

Private Sub defualt_Click()
Call select_color_type(0)
sys_color = "0"
End Sub

Private Sub Del_Click()
frmCancelSerAppointments.Show
End Sub

Private Sub DelDocApp_Click()
frmCancelDocAppointments.Show
End Sub

Private Sub DisplayRoom_Click()
frmRoomDetails.Show
End Sub

Private Sub DisplayWard_Click()
frmWardDetails.Show
End Sub

Private Sub DisSer_Click()
frmService.Show
End Sub

Private Sub DocApp_Click()
frmAddDocAppointments.Show
End Sub

Private Sub DocSalcal_Click()
frm_app_count.Show
End Sub

Private Sub docsched_Click()
frmDoctorSchedule.Show
End Sub



Private Sub editPatient_Click()
frmInPatientBill.Show
End Sub

Private Sub EmpSal_Click()
frm_add_salary_info.Show
End Sub

Public Sub Exit_Click()
AppState = 1
Unload MDIMain

'If MsgBox("Are you sure ?", vbQuestion + vbYesNo, "Confirm Quit Application") = vbYes Then
   'End
'End If

End Sub

Private Sub hosdetails_Click()
frmCompany.Show
End Sub

Private Sub IAddNewPay_Click()
frmIPBillPayments.Show
End Sub

Private Sub Invoice_Click()
frmOrder.Show
End Sub

Private Sub IPDischarge_Click()
frmIPDischarge.Show
End Sub

Private Sub IPDocvisits_Click()
frmInPatientDocReport.Show
End Sub

Private Sub IPMEdIssue_Click()
frmInPatientMedReport.Show
End Sub

Private Sub IPMedServices_Click()
frmInPatientServiceReport.Show
End Sub

Private Sub IViewBill_Click()
frmIPBill.Show
End Sub

Private Sub LightB_Click()
Call select_color_type(5)
sys_color = "5"
End Sub

Private Sub lightV_Click()
Call select_color_type(4)
sys_color = "4"
End Sub

Public Sub logoff_Click()
If MsgBox("Are you sure you want to Log Off the system ?", vbYesNo + vbQuestion, "Log off") = vbYes Then
    AppState = 0
    Unload Me
    frmLogin.Show
End If

End Sub

Private Sub macgrey_Click()
Call select_color_type(1)
sys_color = "1"
End Sub

Private Sub MDIForm_Activate()
Load frmSideBar
Call disMenu
End Sub



Private Sub MDIForm_Load()

original_menu_color = GetSysColor(4)
original_buttonface_color = GetSysColor(15)
original_buttonshadow_color = GetSysColor(16)
original_buttonhighlight_color = GetSysColor(20)
'Set the system color
Call select_color_type(Val(sys_color))


Me.WindowState = vbMaximized
Load frmSideBar

StatusBar1.Panels(2).Text = Date

'Disable Employee Management Section (Not Completed Yet)
Call disMenu


End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'Restore the orignal system color
On Error Resume Next

If AppState = 0 Then
AppState = 1
New_System_Color.SelectColor(4) = original_menu_color
New_System_Color.SelectColor(15) = original_buttonface_color
New_System_Color.SelectColor(16) = original_buttonshadow_color
New_System_Color.SelectColor(20) = original_buttonhighlight_color
Call change_system_color
   
Cancel = 0
Exit Sub
End If



'If MsgBox("Are you sure ?", vbQuestion + vbYesNo, "Confirm Quit Application") = vbYes Then
New_System_Color.SelectColor(4) = original_menu_color
New_System_Color.SelectColor(15) = original_buttonface_color
New_System_Color.SelectColor(16) = original_buttonshadow_color
New_System_Color.SelectColor(20) = original_buttonhighlight_color


Call change_system_color
' Close the database connection on exit
cnPatients.Close
'End
Cancel = 0
   
'Else
   'Cancel = 1
'End If

End Sub



Private Sub MedSerShedrpt_Click()
frmservicesreports.Show
End Sub

Private Sub PInvoice_Click()
FrmPurchases.Show
End Sub

Private Sub Products_Click()
FrmProducts.Show
End Sub

Private Sub RegIPatient_Click()
frmAdmissionDetails.Show
End Sub

Private Sub register1_Click()
frmAppRegister.Show
End Sub



Private Sub RoomRpt_Click()
frmbedwardreport.Show
End Sub

Private Sub rptCustomers_Click()
frmCustomerReport.Show
End Sub

Private Sub rptEmployee_Click()

If MsgBox("This section is currently Under Construction" & vbCrLf & "Do you want to Continue ?", vbQuestion + vbYesNo, "Crystal Employee Management System") = vbYes Then
    frmemployeereports.Show
End If

End Sub

Private Sub rptLogs_Click()
frmLogReport.Show
End Sub

Private Sub rptPurchase_Click()
frmPurchasesReport.Show
End Sub

Private Sub rptSales_Click()
frmSalesReport.Show
End Sub

Private Sub salesbycus_Click()
frmCusSalesReport.Show
End Sub

Private Sub SearchDocApp_Click()
frmViewDoctorAppointments.Show
End Sub

Private Sub SearchSerApp_Click()
frmViewServiceAppointments.Show
End Sub

Private Sub sersched_Click()
frmServiceSchedule.Show
End Sub

Private Sub showempMenu_Click()




If showempMenu.Checked = True Then
    showempMenu.Checked = False
ElseIf showempMenu.Checked = False Then
    showempMenu.Checked = True
End If


If showempMenu.Checked = True Then
    frm_employee.Show
    
ElseIf showempMenu.Checked = False Then
    frm_employee.Hide
End If








End Sub

Private Sub showNavigation_Click()

If showNavigation.Checked = True Then
    showNavigation.Checked = False
ElseIf showNavigation.Checked = False Then
    showNavigation.Checked = True
End If


If showNavigation.Checked = True Then
    FrmNavigation.Show
    
ElseIf showNavigation.Checked = False Then
    FrmNavigation.Hide
End If

End Sub

Private Sub sidebar_Click()

If sidebar.Checked = True Then
    sidebar.Checked = False
ElseIf sidebar.Checked = False Then
    sidebar.Checked = True
End If


If sidebar.Checked = True Then
    frmSideBar.Show
    
ElseIf sidebar.Checked = False Then
    frmSideBar.Hide
End If




End Sub

Private Sub Suppliers_Click()
FrnSuppliers.Show
End Sub

Private Sub Timer1_Timer()
StatusBar1.Panels(1).Text = Time
End Sub

Private Sub Timer2_Timer()

If appRegistered = False Then
    If DateAdd("n", 30, LogTime) = Time Then
        MsgBox "30 Minutes Trial Period Over" & vbCrLf & "Please register the program", vbInformation
        Unload MDIMain
    End If
End If



End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)


Select Case Button
    Case "Out Patient"
        frmAddOutPatientDetails.Show
        
    Case "In Patient"
        frmInPatientDetails.Show
        
    Case "Admission"
        frmAdmissionDetails.Show
End Select


End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

Select Case ButtonMenu
    Case "Doctor Appointment"
    frmAddDocAppointments.Show
    
    Case "Medical Appointment"
    frmAddSerAppointments.Show
    
    Case "In Patient Bill Payments"
        frmIPBillPayments.Show
    Case "Doctor Appointment Payments"
        frmOPBillPayments.Show
    Case "Service Appointment Payments"
        frmOPSerBillPayments.Show
End Select

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button
    Case "Customers"
        FrmCustomers.Show
        
    Case "Sales Invoice"
        frmOrder.Show
        
    
End Select


End Sub

Private Sub ViewDischarge_Click()
frmDisplayIPDischarge.Show
End Sub

Private Sub ViewDoc_Click()
frmDoctorDetails.Show
End Sub

Private Sub ViewDocApp_Click()
frmViewDoctorAppointments.Show
End Sub

Private Sub ViewDocVisits_Click()
frmDoctorVisit.Show
End Sub

Private Sub ViewMedDetails_Click()
frmInPatientOrders.Show
End Sub

Private Sub ViewOBill_Click()
frmOPBillPayments.Show
End Sub

Private Sub ViewOPatient_Click()
frmDisplayOutPatient.Show
End Sub

Private Sub ViewSerApp_Click()
frmViewServiceAppointments.Show
End Sub

Private Sub viewOPHistory_Click()
frmDisplayOPHistory.Show
End Sub

Private Sub ViewSerDetails_Click()
frmIPServiceDetails.Show
End Sub


Private Sub wardrpt_Click()
frmbedwardreport.Show
End Sub

Private Sub winclassic_Click()
Call select_color_type(6)
sys_color = "6"
End Sub

Private Sub xpblue_Click()
Call select_color_type(2)
sys_color = "2"
End Sub
