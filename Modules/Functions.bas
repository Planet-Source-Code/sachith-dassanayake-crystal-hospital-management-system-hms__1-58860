Attribute VB_Name = "Functions"
Option Explicit

Public PvtDocID As String
Public Xtwips As Integer, Ytwips As Integer
Public Xpixels As Integer, Ypixels As Integer

Type FRMSIZE
   Height As Long
   Width As Long
End Type

Public RePosForm As Boolean
Public DoResize As Boolean
Dim MyForm As FRMSIZE
Dim DesignX As Integer
Dim DesignY As Integer
Dim ScaleFactorX As Single, ScaleFactorY As Single
Public Sub SetInitialCaption(Cap As String, Spaces As Integer, FormName As Form)
    FormName.Caption = Space(Spaces)
    FormName.Caption = FormName.Caption + Cap
End Sub

Public Sub ScrollTitle(Cap As String, Spaces As Integer, FormName As Form)
    If Not FormName.Caption = "" Then
        FormName.Caption = Right(FormName.Caption, (Len(FormName.Caption) - 1))
    Else
        Call SetInitialCaption(Cap, Spaces, FormName)
    End If
End Sub
' This function is used to create Unique ID's
Public Function UID(mLen As Integer, mPrefix As String) As String
    
    Dim mStr As String, i As Integer, j As Integer, mTable() As String * 1
    ReDim mTable(1 To 61)
    mTable(1) = "1": mTable(2) = "2": mTable(3) = "3": mTable(4) = "4"
    mTable(5) = "5": mTable(6) = "6": mTable(7) = "7": mTable(8) = "8"
    mTable(9) = "9": mTable(10) = "0"
    mTable(11) = "a": mTable(12) = "b": mTable(13) = "c": mTable(14) = "d"
    mTable(15) = "e": mTable(16) = "f": mTable(17) = "g": mTable(18) = "h"
    mTable(19) = "i": mTable(20) = "j": mTable(21) = "k": mTable(22) = "l"
    mTable(23) = "m": mTable(24) = "n": mTable(25) = "o": mTable(26) = "p"
    mTable(27) = "q": mTable(28) = "r": mTable(29) = "s": mTable(30) = "t"
    mTable(31) = "u": mTable(32) = "v": mTable(33) = "w": mTable(34) = "x"
    mTable(35) = "y": mTable(36) = "z"
    mTable(37) = "A": mTable(38) = "B": mTable(39) = "C": mTable(40) = "D"
    mTable(41) = "E": mTable(42) = "F": mTable(43) = "G": mTable(44) = "H":
    mTable(45) = "I": mTable(46) = "J": mTable(47) = "K": mTable(48) = "L"
    mTable(49) = "M": mTable(50) = "N": mTable(51) = "O": mTable(52) = "P"
    mTable(52) = "Q": mTable(53) = "R": mTable(54) = "S": mTable(55) = "T"
    mTable(56) = "U": mTable(57) = "V": mTable(58) = "W": mTable(59) = "X":
    mTable(60) = "Y": mTable(61) = "Z":
    mStr = mPrefix
    For i = 1 To mLen
    
    For j = 0 To 10
        DoEvents
        Next j
        Randomize
        mStr = mStr & mTable(Int((60) * Rnd + 1))
    Next i
    
        UID = mStr
        
End Function
    
Public Sub SizeColumns(ByVal flx As MSFlexGrid, frm As Form)

Dim max_wid As Single
Dim wid As Single
Dim max_row As Integer
Dim r As Integer
Dim c As Integer

    max_row = flx.Rows - 1
    For c = 0 To flx.Cols - 1
        max_wid = 0
        For r = 0 To max_row
        'wid = TextWidth(flx.TextMatrix(r, c))
        wid = frm.TextWidth(flx.TextMatrix(r, c))
            If max_wid < wid Then max_wid = wid
        Next r
        flx.ColWidth(c) = max_wid + wid
    Next c
 
End Sub
Public Sub SizeColumns1(ByVal flx As MSHFlexGrid, frm As Form)

Dim max_wid As Single
Dim wid As Single
Dim max_row As Integer
Dim r As Integer
Dim c As Integer

    max_row = flx.Rows - 1
    For c = 0 To flx.Cols - 1
        max_wid = 0
        For r = 0 To max_row
        'wid = TextWidth(flx.TextMatrix(r, c))
        wid = frm.TextWidth(flx.TextMatrix(r, c))
            If max_wid < wid Then max_wid = wid
        Next r
        flx.ColWidth(c) = max_wid + wid
    Next c
End Sub

Public Sub SizeColumnHeaders(ByVal flx As MSFlexGrid, frm As Form)
Dim max_wid As Single
Dim wid As Single
Dim max_row As Integer
Dim r As Integer
Dim c As Integer

max_wid = 0
        
    max_row = flx.Rows - 1
    For c = 0 To flx.Cols - 1
        'wid = TextWidth(flx.TextMatrix(0, c))
         wid = frm.TextWidth(flx.TextMatrix(r, c))
        If max_wid < wid Then
            max_wid = wid
        End If
        
        flx.ColWidth(c) = max_wid + wid
    Next c
End Sub

Public Function SQLDate(ConvertDate As Date) As String
    SQLDate = Format(ConvertDate, "mm/dd/yyyy")
End Function


Public Function DataEntryValidation(Key As Integer, Param As String) As Integer
    'If BckSpace then allow
    If Key = 8 Then DataEntryValidation = Key: Exit Function
    'Enforce only Digits

    Select Case Param
        Case "Num"
        If Key < Asc("0") Or Key > Asc("9") Then
            DataEntryValidation = 0
        Else
            DataEntryValidation = Key
        End If
        
        Case "Amt"
        If Key < Asc("0") Or Key > Asc("9") Then

            If Key <> Asc(".") Then
                DataEntryValidation = 0
            Else
                DataEntryValidation = Key
            End If
        Else
            DataEntryValidation = Key
        End If
        
        Case "Chr"
        Key = Asc(UCase(Chr(Key)))
        If Key = 8 Or Key = 32 Then DataEntryValidation = Key: Exit Function
        If Key < Asc("A") Or Key > Asc("Z") Then
            DataEntryValidation = 0
        Else
            DataEntryValidation = Asc(UCase(Chr(Key)))
        End If
        
        Case "Cbo"
        DataEntryValidation = 0
        
 
        Case Else
        DataEntryValidation = Asc(UCase(Chr(Key)))
        
    End Select
End Function


Public Function Encrypt(StringToEncrypt As String, Optional AlphaEncoding As Boolean = False) As String
    On Error GoTo ErrorHandler
    Dim Char As String
    Dim i As Integer
    Encrypt = ""
    


    For i = 1 To Len(StringToEncrypt)
        Char = Asc(MID(StringToEncrypt, i, 1))
        Encrypt = Encrypt & Len(Char) & Char
    Next i
    

    If AlphaEncoding Then
        
        StringToEncrypt = Encrypt
        Encrypt = ""


        For i = 1 To Len(StringToEncrypt)
            Encrypt = Encrypt & Chr(MID(StringToEncrypt, i, 1) + 147)
        Next i
    End If
    Exit Function
ErrorHandler:
    Encrypt = "Error encrypting string"
End Function

Public Function Decrypt(StringToDecrypt As String, Optional AlphaDecoding As Boolean = False) As String
    On Error GoTo ErrorHandler
    Dim CharCode As String
    Dim CharPos As Integer
    Dim Char As String
    Dim i As Integer
    

    If AlphaDecoding Then
        
        Decrypt = StringToDecrypt
        StringToDecrypt = ""


        For i = 1 To Len(Decrypt)
            StringToDecrypt = StringToDecrypt & (Asc(MID(Decrypt, i, 1)) - 147)
        Next i
    End If
    
   Decrypt = ""
    

    Do
        
        CharPos = Left(StringToDecrypt, 1)
        StringToDecrypt = MID(StringToDecrypt, 2)
        CharCode = Left(StringToDecrypt, CharPos)
       StringToDecrypt = MID(StringToDecrypt, Len(CharCode) + 1)
        Decrypt = Decrypt & Chr(CharCode)
    Loop Until StringToDecrypt = ""
    Exit Function
ErrorHandler:
    Decrypt = "Error decrypting string"
End Function

Public Function Enc(Epass As String) As String

Epass = StrReverse(Epass)
Enc = EID(12, Epass) & EID(8)



End Function
Public Function Dec(Dpass As String) As String
On Error Resume Next
Dpass = StrReverse(Dpass)
   

Dec = Left(Dpass, Len(Dpass) - 12)
Dec = Right(Dec, Len(Dec) - 8)

End Function
Public Function EID(mLen As Integer, Optional mPrefix As String) As String
    
    Dim mStr As String, i As Integer, j As Integer, mTable() As String * 1
    ReDim mTable(1 To 61)
    mTable(1) = "1": mTable(2) = "2": mTable(3) = "3": mTable(4) = "4"
    mTable(5) = "5": mTable(6) = "6": mTable(7) = "7": mTable(8) = "8"
    mTable(9) = "9": mTable(10) = "0"
    mTable(11) = "a": mTable(12) = "b": mTable(13) = "c": mTable(14) = "d"
    mTable(15) = "e": mTable(16) = "f": mTable(17) = "g": mTable(18) = "h"
    mTable(19) = "i": mTable(20) = "j": mTable(21) = "k": mTable(22) = "l"
    mTable(23) = "m": mTable(24) = "n": mTable(25) = "o": mTable(26) = "p"
    mTable(27) = "q": mTable(28) = "r": mTable(29) = "s": mTable(30) = "t"
    mTable(31) = "u": mTable(32) = "v": mTable(33) = "w": mTable(34) = "x"
    mTable(35) = "y": mTable(36) = "z"
    mTable(37) = "A": mTable(38) = "B": mTable(39) = "C": mTable(40) = "D"
    mTable(41) = "E": mTable(42) = "F": mTable(43) = "G": mTable(44) = "H":
    mTable(45) = "I": mTable(46) = "J": mTable(47) = "K": mTable(48) = "L"
    mTable(49) = "M": mTable(50) = "N": mTable(51) = "O": mTable(52) = "P"
    mTable(52) = "Q": mTable(53) = "R": mTable(54) = "S": mTable(55) = "T"
    mTable(56) = "U": mTable(57) = "V": mTable(58) = "W": mTable(59) = "X":
    mTable(60) = "Y": mTable(61) = "Z":
    'mStr = mPrefix


    For i = 1 To mLen


        For j = 0 To 10


            DoEvents
            Next j
            Randomize
            mStr = mStr & mTable(Int((60) * Rnd + 1))
        Next i
        EID = mStr & mPrefix
End Function

Public Sub DisableMenu()

On Error Resume Next

Dim ctl As Control
For Each ctl In MDIMain.Controls
    
    If TypeOf ctl Is Timer Then
    Else
    ctl.Enabled = False
    End If
Next

End Sub
Public Sub EnableMenu()
On Error Resume Next

Dim ctl As Control
For Each ctl In MDIMain.Controls
    ctl.Enabled = True
Next
End Sub



