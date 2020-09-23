VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.4#0"; "AResize.ocx"
Begin VB.Form frmIPServiceDetails 
   BackColor       =   &H00FF8080&
   Caption         =   "In Patient Services Details"
   ClientHeight    =   6030
   ClientLeft      =   1110
   ClientTop       =   450
   ClientWidth     =   11580
   Icon            =   "frmIPServiceDetails.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   11580
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   10200
      Picture         =   "frmIPServiceDetails.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   1080
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      DragIcon        =   "frmIPServiceDetails.frx":0DCE
      Height          =   3720
      Left            =   180
      TabIndex        =   1
      Top             =   1260
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   6562
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Rows            =   4
      Cols            =   11
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      MergeCells      =   4
      AllowUserResizing=   1
      FormatString    =   "|In Patient Service ID|Service ID|In Patient ID|Admission ID|Bill Date|Service Date|Service Time|Service Charge|Discount|Total"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
      _Band(0).GridLinesBand=   1
      _Band(0).GridLineWidthBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
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
      FormHeightDT    =   6540
      FormWidthDT     =   11700
      FormScaleHeightDT=   6030
      FormScaleWidthDT=   11580
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF8080&
      Caption         =   "IN PATIENT SERVICE DETAILS"
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
      Left            =   3000
      TabIndex        =   2
      Top             =   360
      Width           =   6105
   End
End
Attribute VB_Name = "frmIPServiceDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MARGIN_SIZE = 60      ' in Twips
' variables for data binding
Private datPrimaryRS As ADODB.Recordset

' variables for enabling column sort
Private m_iSortCol As Integer
Private m_iSortType As Integer

' variables for column dragging
Private m_bDragOK As Boolean
Private m_iDragCol As Integer
Private xdn As Integer, ydn As Integer

Private Sub Form_Load()
    Call Functions.DisableMenu
    Dim sConnect As String
    Dim sSQL As String
    ' set strings
    sSQL = "select InPatient_ServiceID,Hospital_Service_ID,InPatientID,AdmissionID,Bill_Date,Service_Date,Service_Time,Service_Charge,Discount,Total from InPatient_Services"

    ' open connection
    

    ' create a recordset using the provided collection
    Set datPrimaryRS = New Recordset
    datPrimaryRS.CursorLocation = adUseClient
    datPrimaryRS.Open sSQL, cnPatients, adOpenForwardOnly, adLockReadOnly

    Set MSHFlexGrid1.DataSource = datPrimaryRS

    With MSHFlexGrid1

        .Redraw = False
        ' set grid's column widths
        .ColWidth(0) = 300
        .ColWidth(1) = -1
        .ColWidth(2) = -1
        .ColWidth(3) = -1
        .ColWidth(4) = -1
        .ColWidth(5) = -1
        .ColWidth(6) = -1
        .ColWidth(7) = -1
        .ColWidth(8) = -1
        .ColWidth(9) = -1
        .ColWidth(10) = -1

        ' set grid's style
        .AllowBigSelection = True
        .FillStyle = flexFillRepeat

        ' make header bold
        .Row = 0
        .Col = 0
        .RowSel = .FixedRows - 1
        .ColSel = .Cols - 1
        .CellFontBold = True

        .AllowBigSelection = False
        .FillStyle = flexFillSingle
        
        .Redraw = True
        
        MSHFlexGrid1.TextMatrix(0, 1) = "In Patient ServiceID"
        MSHFlexGrid1.TextMatrix(0, 2) = "Service ID"
        MSHFlexGrid1.TextMatrix(0, 3) = "In Patient ID"
        MSHFlexGrid1.TextMatrix(0, 4) = "Admission ID"
        MSHFlexGrid1.TextMatrix(0, 5) = "Bill Date"
        MSHFlexGrid1.TextMatrix(0, 6) = "Service Date"
        MSHFlexGrid1.TextMatrix(0, 7) = "Service Time"
        MSHFlexGrid1.TextMatrix(0, 8) = "Service Charge"
        MSHFlexGrid1.TextMatrix(0, 9) = "Discount"
        MSHFlexGrid1.TextMatrix(0, 10) = "Total"
        
        Functions.SizeColumns1 MSHFlexGrid1, Me
        
        MSHFlexGrid1.ColWidth(8) = 1000
        MSHFlexGrid1.ColWidth(9) = 1000
        MSHFlexGrid1.ColWidth(4) = 1800
        
        

    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Call Functions.EnableMenu
End Sub

Private Sub MSHFlexGrid1_DragDrop(Source As Control, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' code in grid's DragDrop, MouseDown, MouseMove, and MouseUp events enables column dragging
'-------------------------------------------------------------------------------------------

    If m_iDragCol = -1 Then Exit Sub    ' we weren't dragging
    If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub
    If MSHFlexGrid1.FixedCols = 1 And MSHFlexGrid1.MouseCol = 0 Then Exit Sub

    With MSHFlexGrid1
        .Redraw = False
        .ColPosition(m_iDragCol) = .MouseCol
        .Redraw = True
    End With

End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' code in grid's DragDrop, MouseDown, MouseMove, and MouseUp events enables column dragging
'-------------------------------------------------------------------------------------------

    If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub
    If MSHFlexGrid1.MouseCol = 0 And MSHFlexGrid1.FixedCols = 1 Then Exit Sub

    xdn = X
    ydn = Y
    m_iDragCol = -1     ' clear drag flag
    m_bDragOK = True

End Sub

Private Sub MSHFlexGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' code in grid's DragDrop, MouseDown, MouseMove, and MouseUp events enables column dragging
'-------------------------------------------------------------------------------------------

    ' test to see if we should start drag
    If Not m_bDragOK Then Exit Sub
    If Button <> 1 Then Exit Sub                        ' wrong button
    If m_iDragCol <> -1 Then Exit Sub                   ' already dragging
    If Abs(xdn - X) + Abs(ydn - Y) < 50 Then Exit Sub   ' didn't move enough yet
    If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub         ' must drag header

    ' if got to here then start the drag
    m_iDragCol = MSHFlexGrid1.MouseCol
    MSHFlexGrid1.Drag vbBeginDrag

End Sub

Private Sub MSHFlexGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' code in grid's DragDrop, MouseDown, MouseMove, and MouseUp events enables column dragging
'-------------------------------------------------------------------------------------------

    m_bDragOK = False

End Sub

Private Sub MSHFlexGrid1_DblClick()
'-------------------------------------------------------------------------------------------
' code in grid's DblClick event enables column sorting
'-------------------------------------------------------------------------------------------

    Dim i As Integer

    ' sort only when a fixed row is clicked
    If MSHFlexGrid1.MouseRow >= MSHFlexGrid1.FixedRows Then Exit Sub

    i = m_iSortCol                  ' save old column
    m_iSortCol = MSHFlexGrid1.Col   ' set new column

    ' increment sort type
    If i <> m_iSortCol Then
        ' if clicking on a new column, start with ascending sort
        m_iSortType = 1
    Else
        ' if clicking on the same column, toggle between ascending and descending sort
        m_iSortType = m_iSortType + 1
    If m_iSortType = 3 Then m_iSortType = 1
    End If

    DoColumnSort

End Sub

Sub DoColumnSort()
'-------------------------------------------------------------------------------------------
' does Exchange-type sort on column m_iSortCol
'-------------------------------------------------------------------------------------------

    With MSHFlexGrid1
        .Redraw = False
        .Row = 1
        .RowSel = .Rows - 1
        .Col = m_iSortCol
        .Sort = m_iSortType
        .Redraw = True
    End With

End Sub

Private Sub Form_Resize()

    Dim sngButtonTop As Single
    Dim sngScaleWidth As Single
    Dim sngScaleHeight As Single

    On Error GoTo Form_Resize_Error
    With Me
        sngScaleWidth = .ScaleWidth
        sngScaleHeight = .ScaleHeight

        ' move Close button to the lower right corner
        With .cmdClose
                sngButtonTop = sngScaleHeight - (.Height + MARGIN_SIZE)
                .Move sngScaleWidth - (.Width + MARGIN_SIZE), sngButtonTop
        End With

        .MSHFlexGrid1.Move MARGIN_SIZE, _
            MARGIN_SIZE, _
            sngScaleWidth - (2 * MARGIN_SIZE), _
            sngButtonTop - (2 * MARGIN_SIZE)

    End With
    Exit Sub

Form_Resize_Error:
    ' avoid error on negative values
    Resume Next

End Sub
Private Sub cmdClose_Click()

    Unload Me

End Sub


