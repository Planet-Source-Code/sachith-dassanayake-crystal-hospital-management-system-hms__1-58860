VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDisplayOutPatient 
   BackColor       =   &H80000003&
   Caption         =   "Out Patient Details"
   ClientHeight    =   5160
   ClientLeft      =   1110
   ClientTop       =   450
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5160
   ScaleWidth      =   7200
   Begin VB.ComboBox cmbSort 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   300
      Left            =   5880
      TabIndex        =   0
      Top             =   4560
      Width           =   1080
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      DragIcon        =   "frmDisplayOutPatient.frx":0000
      Height          =   4320
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   7620
      _Version        =   393216
      BackColor       =   -2147483624
      ForeColor       =   0
      Rows            =   19
      Cols            =   9
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      MergeCells      =   4
      AllowUserResizing=   1
      FormatString    =   "|Patient ID|First Name|Last Name|Gender|Address|Telephone|Status|Notes"
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
      _Band(0).Cols   =   9
      _Band(0).GridLineWidthBand=   1
      _Band(0).TextStyleBand=   0
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000003&
      Caption         =   "Sort List by:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
End
Attribute VB_Name = "frmDisplayOutPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MARGIN_SIZE = 60      ' in Twips
' variables for data binding
Private datPrimaryRS As ADODB.Recordset

' variables for column dragging
Private m_bDragOK As Boolean
Private m_iDragCol As Integer
Private xdn As Integer, ydn As Integer
Private m_iSortCol As Integer
Private m_iSortType As Integer





Private Sub cmbSort_Click()


Dim i As Integer

    ' sort only when a fixed row is clicked
    'If MSHFlexGrid1.MouseRow >= MSHFlexGrid1.FixedRows Then Exit Sub

    i = m_iSortCol ' save old column
    'm_iSortCol = MSHFlexGrid1.Col   ' set new column
    m_iSortCol = cmbSort.ListIndex + 1
    

    Debug.Print MSHFlexGrid1.Col
    
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

Private Sub Form_Load()

    Dim sConnect As String
    Dim sSQL As String
    'Dim dfwConn As ADODB.Connection
    Dim i As Integer

    ' set strings
    'sConnect = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;User ID=Admin;Data Source=D:\test\Appoinment1.mdb;Mode=Share Deny None;Extended Properties=';COUNTRY=0;CP=1252;LANGID=0x0409';Jet OLEDB:System database='';Jet OLEDB:Registry Path='';Jet OLEDB:Database Password='';Jet OLEDB:Global Partial Bulk Ops=2"
    sSQL = "select Patient_ID,First_Name,Last_Name,Gender,Address,Telephone,Status,Notes from Patient_Details"

    ' open connection
    'Set dfwConn = New Connection
    'dfwConn.Open sConnect

    ' create a recordset using the provided collection
    Set datPrimaryRS = New Recordset
    datPrimaryRS.CursorLocation = adUseClient
    datPrimaryRS.Open sSQL, cnPatients, adOpenForwardOnly, adLockReadOnly

    Set MSHFlexGrid1.DataSource = datPrimaryRS
    
    

    With MSHFlexGrid1

        .Redraw = False
        ' set grid's column widths
        .ColWidth(0) = 300
        .ColWidth(1) = 1080
        .ColWidth(2) = 1245
        .ColWidth(3) = 1185
        .ColWidth(4) = -1
        .ColWidth(5) = 1560
        .ColWidth(6) = 1335
        .ColWidth(7) = -1
        .ColWidth(8) = -1
        

        .TextMatrix(0, 1) = "Patient ID"
        .TextMatrix(0, 2) = "First Name"
        .TextMatrix(0, 3) = "Last Name"
        .TextMatrix(0, 4) = "Gender"
        .TextMatrix(0, 5) = "Address"
        .TextMatrix(0, 6) = "Telephone"
        .TextMatrix(0, 7) = "Status"
        .TextMatrix(0, 8) = "Notes"

        
        Functions.SizeColumns1 MSHFlexGrid1, Me
        ' set grid's column merging and sorting
        For i = 0 To .Cols - 2
            .MergeCol(i) = True
            cmbSort.AddItem datPrimaryRS(i).Name, i
        Next i


        .Sort = flexSortGenericAscending

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

    End With

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
        DoSort
        .Redraw = True
    End With

End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
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

Private Sub MSHFlexGrid1_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
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

Private Sub MSHFlexGrid1_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' code in grid's DragDrop, MouseDown, MouseMove, and MouseUp events enables column dragging
'-------------------------------------------------------------------------------------------

    m_bDragOK = False

End Sub

Sub DoSort()

    With MSHFlexGrid1
        .Redraw = False
        .Col = 0
        .Row = 1
        .RowSel = .Rows - 1
        .Sort = flexSortGenericAscending
        .Redraw = True
    End With

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
        
        
        sngButtonTop = sngScaleHeight - (cmbSort.Height + MARGIN_SIZE)
        cmbSort.Move 0 + (cmbSort.Width + MARGIN_SIZE), sngButtonTop
        sngButtonTop = sngScaleHeight - (Label1.Height + MARGIN_SIZE)
        Label1.Move 0 + Label1.Width - 150, sngButtonTop
        

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


