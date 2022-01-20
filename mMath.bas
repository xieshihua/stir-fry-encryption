Attribute VB_Name = "mMath"
Option Compare Database
Option Explicit

' # File: mMath
' # Created by: Robert Xie
' # Updated: 2022-01-19
' # Requires: Microsoft VBScript Regular Expressions 5.5
' # mMath module includes supporting functions and subs for VBA application devolopment. _
    Copyright (C) <2022>  <Robert Xie> _
    This program is free software: you can redistribute it and/or modify _
    it under the terms of the GNU General Public License as published by _
    the Free Software Foundation, either version 3 of the License, or _
    (at your option) any later version. _
    This program is distributed in the hope that it will be useful, _
    but WITHOUT ANY WARRANTY; without even the implied warranty of _
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the _
    GNU General Public License for more details. _
    You should have received a copy of the GNU General Public License _
    along with this program.  If not, see <http://www.gnu.org/licenses/>.

Const ONE_BYTE = 8

Public Function Help(Optional stFunctionName As String = "?") As String
    Dim stF As String
    Dim tmp As String
    
    stF = UCase(Nz(stFunctionName, "?"))
    stF = IIf(stF = "", "?", stF)
    Select Case stF
    Case "HELP"
        tmp = "Usage: Help(function_name)" & vbCrLf & _
            "       Help(?) for a list of functions."
    Case "?"
        tmp = "The module contains following functions:" & vbCrLf & _
            "Help" & vbCrLf & _
            "" & vbCrLf & _
            "Usage: Help(function_name)"
    Case Else
        tmp = stFunctionName & " does not exit in this module."
    End Select
    
    Help = tmp
End Function

' Function: ArrayMerge
' Takes: two arrays
' Returns: a merged array with the first element as the merged field, and 2nd element as source indicator:
'       - "=", exist in both array
'       - "+", exist only in the 2nd array
'       - "-", exist only in the 1st array
'       - "0", blank string
Public Function ArrayMerge(ByRef ArrayPrime() As Variant, ArraySecondary() As Variant) As Variant()
On Error GoTo Err_ArrayMerge
    Dim primeL As Integer, primeU As Integer
    Dim secondL As Integer, secondU As Integer
    Dim i As Integer, j As Integer, k As Integer
    Dim arTmp() As Variant
    Dim arResult() As Variant
    
    primeL = LBound(ArrayPrime): primeU = UBound(ArrayPrime)
    secondL = LBound(ArraySecondary): secondU = UBound(ArraySecondary)
    
    ArraySort ArrayPrime, primeU + 1
    ArraySort ArraySecondary, secondU + 1
    
    ReDim arTmp(primeU + secondU + 1)
    arTmp(0) = Array("", "")
    i = 0: j = 0: k = 0
    While i <= primeU Or j <= secondU
        If i <= primeU And j <= secondU Then
            If ArrayPrime(i) = ArraySecondary(j) Then
                If Nz(ArrayPrime(i), "") = "" Then
                    arTmp(k) = Array(ArrayPrime(i), "0")
                Else
                    arTmp(k) = Array(ArrayPrime(i), "=")
                End If
                i = i + 1: j = j + 1
            ElseIf ArrayPrime(i) > ArraySecondary(j) Then
                arTmp(k) = Array(ArraySecondary(j), "+")
                j = j + 1
            Else
                arTmp(k) = Array(ArrayPrime(i), "-")
                i = i + 1
            End If
            
            If arTmp(k)(0) <> "" Then
                k = k + 1
            End If
        ElseIf i <= primeU Then
            arTmp(k) = Array(ArrayPrime(i), "-")
            i = i + 1
            k = k + 1
        Else
            arTmp(k) = Array(ArraySecondary(j), "+")
            j = j + 1
            k = k + 1
        End If
    Wend
    
    ReDim arResult(IIf(k > 0, k - 1, 0))
    For i = 0 To k - 1
        arResult(i) = arTmp(i)
    Next i
Exit_ArrayMerge:
    ArrayMerge = arResult
    Exit Function
Err_ArrayMerge:
    ReDim arResult(0)
    arResult(0) = ""
    MsgBox Err.Description, vbInformation, "mUtilities.ArrayMerge"
    Resume Exit_ArrayMerge
End Function

' Sort array of associated values passed in through ar descendently by using bubble sort.
Public Sub ArraySort(ByRef ar() As Variant, size As Integer)
On Error GoTo Err_ArraySort
    Dim i As Integer, j As Integer
    Dim tmp As Variant
    size = IIf(size > UBound(ar) + 1, UBound(ar) + 1, size)
    For i = size - 1 To 1 Step -1
        For j = i - 1 To 0 Step -1
            If Nz(ar(i), "") < Nz(ar(j), "") Then
                tmp = ar(i)
                ar(i) = ar(j)
                ar(j) = tmp
            End If
        Next j
    Next i
Exit_ArraySort:
    Exit Sub
Err_ArraySort:
    MsgBox Err.Description, vbInformation, "Math.ArraySort"
    Resume Exit_ArraySort
End Sub

' Sort array of associated values passed in through ar by using bubble sort.
Public Sub ArraySort_Associated(ByRef ar() As Variant, size As Integer, _
                        Optional sort_by_column As Integer = 1, Optional descending As Boolean = False)
On Error GoTo Err_ArraySort_Associated
    Dim i As Integer, j As Integer
    Dim tmp As Variant
    Dim lb As Integer
    
    lb = LBound(ar)
    size = Min(size, UBound(ar) + 1)
    
    For i = size - 1 To lb + 1 Step -1
        For j = lb + 1 To i
            If (ar(j - 1)(sort_by_column) < ar(j)(sort_by_column)) = descending Then
                tmp = ar(j)
                ar(j) = ar(j - 1)
                ar(j - 1) = tmp
            End If
        Next j
    Next i
Exit_ArraySort_Associated:
    Exit Sub
Err_ArraySort_Associated:
    MsgBox Err.Description, vbInformation, "Math.ArraySort_Associated"
    Resume Exit_ArraySort_Associated
End Sub

Public Function ArrayToString(ByRef ar() As Variant) As String
    Dim i As Integer
    Dim tmp As String
    tmp = ""
    For i = LBound(ar) To UBound(ar)
        tmp = IIf(tmp = "", "", tmp & vbCrLf) & Nz(ar(i), "")
    Next i
    ArrayToString = tmp
End Function

' http://answers.microsoft.com/en-us/office/forum/office_2003-access/access-trigonometric-functions/03965421-4da4-4b2c-a023-9337266b27fb?auth=1
' VBA Help: Deraived Math Functions

Public Function ArcSin(x As Double) As Double
    ArcSin = Atn(x / Sqr(-x * x + 1))
End Function

Public Function ArcCos(x As Double) As Double
    ArcCos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
End Function

Public Function Sinh(x As Double) As Double
    Sinh = (Exp(x) - Exp(-1 * x)) / 2
End Function

Public Function ArcSinh(x As Double) As Double
    ArcSinh = Log(x + Sqr(x * x + 1))
End Function

Public Function Cosh(x As Double) As Double
    Cosh = (Exp(x) + Exp(-1 * x)) / 2
End Function

Public Function ArcCosh(x As Double) As Double
    ArcCosh = Log(x + Sqr(x * x - 1))
End Function

Public Function IsLeapYear(intYear As Integer) As Boolean
    IsLeapYear = (intYear Mod 4) = 0 And (intYear Mod 100) <> 0
End Function

Public Function Max(first As Variant, ParamArray Others() As Variant) As Variant
On Error GoTo Err_Max
    Dim i As Integer
    Dim tmp As Variant
    
    tmp = first
    For i = 0 To UBound(Others)
        If Others(i) > tmp Then
            tmp = Others(i)
        End If
    Next i
    
    Max = tmp
Exit_Max:
    Exit Function
Err_Max:
    MsgBox Err.Description, vbInformation, "mMath.Max"
    Max = Null
    Resume Exit_Max
End Function

Public Function Min(first As Variant, ParamArray Others() As Variant) As Variant
On Error GoTo Err_Min
    Dim i As Integer
    Dim tmp As Variant
    
    tmp = first
    For i = 0 To UBound(Others)
        If Others(i) < tmp Then
            tmp = Others(i)
        End If
    Next i
    Min = tmp
Exit_Min:
    Exit Function
Err_Min:
    MsgBox Err.Description, vbInformation, "mMath.Min"
    Min = Null
    Resume Exit_Min
End Function

' MinNoneZero:
' Takes: (1...many) numbers of any type.
' Returns: the smallest none zero value. If all values are zero, it will return zero.
Public Function MinNoneZero(first As Variant, ParamArray Others() As Variant) As Variant
On Error GoTo Err_MinNoneZero
    Dim i As Integer
    Dim tmp As Variant
    
    tmp = first
    For i = 0 To UBound(Others)
        If Others(i) > 0 Then
            If tmp = 0 Then
                tmp = Others(i)
            ElseIf Others(i) < tmp Then
                tmp = Others(i)
            End If
        End If
    Next i
Exit_MinNoneZero:
    MinNoneZero = tmp
    Exit Function
Err_MinNoneZero:
    MsgBox Err.Description, vbInformation, "mMath.MinNoneZero"
    tmp = 0
    Resume Exit_MinNoneZero
End Function

Public Function m_AddMonth(dtDate As Date, intMonth As Integer) As Date
    Dim dtNew As Date
    
    dtNew = DateAdd("m", intMonth, dtDate) + (-1) * m_SignOf(CDbl(intMonth))
    If IsLeapYear(Year(dtNew)) And Month(dtNew) = 2 And intMonth <> 0 Then
        dtNew = dtNew + 1
    End If
    
    m_AddMonth = dtNew
End Function

Public Function m_DateToString(dtDate As Date) As String
    m_DateToString = Year(dtDate) & "-" & IIf(Month(dtDate) < 10, "0", "") & Month(dtDate) & _
        "-" & IIf(Day(dtDate) < 10, "0", "") & Day(dtDate)
End Function

Public Function m_SignOf(number As Double) As Integer
    Select Case number
    Case Is > 0
        m_SignOf = 1
    Case Is = 0
        m_SignOf = 0
    Case Is < 0
        m_SignOf = -1
    End Select
End Function

' https://bytes.com/topic/access/answers/558113-function-pi
Public Function PI() As Double
    PI = 4 * Atn(1)
End Function

Public Function getFiscal(theDate As Variant) As String
On Error GoTo err_getFiscal
    Dim stFiscal As String
    
    stFiscal = "0000"
    'If IsMissing(theDate) Then
        'stFiscal = "Not Defined" 'Error code for the fiscal
    If IsDate(Nz(theDate, 0)) Then
        stFiscal = Year(DateAdd("m", 9, theDate))
    End If
Exit_getFiscal:
    getFiscal = stFiscal
    Exit Function
err_getFiscal:
    MsgBox Err.Description, vbOKOnly, "getFiscal"
    stFiscal = "Error"
    Resume Exit_getFiscal
End Function

Public Function getFiscalInt(theDate As Variant) As Integer
On Error GoTo err_getFiscalInt
    Dim intFiscal As Integer
    
    intFiscal = -1
    
    If IsDate(Nz(theDate, 0)) Then
        intFiscal = Year(DateAdd("m", 9, theDate))
    End If
Exit_getFiscalInt:
    getFiscalInt = intFiscal
    Exit Function
err_getFiscalInt:
    MsgBox Err.Description, vbOKOnly, "getFiscalInt"
    Resume Exit_getFiscalInt
End Function

Public Function GetFiscalQuarter(theDate As Variant) As String
    Dim tmp As String
        
    tmp = "0" 'Error code for the quarter
    If IsDate(Nz(theDate, 0)) Then
        tmp = IIf(Month(theDate) Mod 3 = 0, 0, 1) + Int(Month(DateAdd("m", -3, theDate)) / 3)
    End If
    GetFiscalQuarter = "FQ" & tmp
End Function

Public Function GetYearQuarter(theDate As Variant) As String
    Dim tmp As String
    
    tmp = "0" 'Error code for the quarter
    If IsDate(Nz(theDate, 0)) Then
        tmp = IIf(Month(theDate) Mod 3 = 0, 0, 1) + Int(Month(theDate) / 3)
    End If
    GetYearQuarter = "YQ" & tmp
End Function

Public Function getQuarter(theDate As Variant) As String
    Dim tmp As String
        
    tmp = "0" 'Error code for the quarter
    If IsDate(Nz(theDate, 0)) Then
        tmp = IIf(Month(theDate) Mod 3 = 0, 0, 1) + Int(Month(DateAdd("m", -3, theDate)) / 3)
    End If
    getQuarter = "Q" & tmp
End Function

Public Function getYearDiff(fstDate As Date, sndDate As Date) As Integer
    Dim intMD As Integer
    Dim tmp As Date
    
    ' Sort the dates
    If fstDate > sndDate Then
        tmp = fstDate
        fstDate = sndDate
        sndDate = tmp
    End If
    
    intMD = DateDiff("m", fstDate, sndDate)
    If (Month(fstDate) = Month(sndDate)) And intMD > 11 Then
        intMD = intMD + IIf(DatePart("d", fstDate) > DatePart("d", sndDate), -1, 0)
    End If
    
    getYearDiff = Fix(intMD / 12)
End Function

Public Sub ByteToBinary(nbrByte As Byte, arrayBinary As Variant)
    ' Takes: a byte number and a boolean array of 8.
    ' Returns: the result is returned in the boolean array passed in.
    Dim ub As Integer, i As Byte
    
    ub = UBound(arrayBinary)
    
    If ub > ONE_BYTE - 1 Then
        MsgBox "Subscription " & ub & " is larger than " & ONE_BYTE - 1 & ".", vbInformation, "ByteToBinary"
        Exit Sub
    End If
    
    
    For i = 0 To ub
        arrayBinary(i) = Int(nbrByte / 2 ^ i) Mod 2
    Next i
End Sub

Public Function ByteToBinaryString(nbrByte As Byte) As String
    Dim arBinary(ONE_BYTE - 1) As Boolean
    Dim i As Integer
    Dim tmp As String
    
    ByteToBinary nbrByte, arBinary
    
    tmp = ""
    For i = 0 To ONE_BYTE - 1
        tmp = IIf(arBinary(i), 1, 0) & tmp
    Next i
    
    ByteToBinaryString = tmp
End Function

Public Function ByteArrayToBinaryString(arByteNbr As Variant) As String
    Dim iLB As Long, iUB As Long
    Dim tmp As String
    Dim i As Long
    
    iLB = LBound(arByteNbr)
    iUB = UBound(arByteNbr)
    
    tmp = ""
    
    For i = iLB To iUB
        tmp = tmp & IIf(i = 0, "", ", ") & ByteToBinaryString(CByte(arByteNbr(i)))
    Next i
    
    ByteArrayToBinaryString = tmp
End Function

Public Function StringToByteArray(strInput As String) As Variant
    ' arByteNbr must be defined as arByte()
    Dim arByteBlock() As Byte
    Dim iLen As Long
    Dim i As Long
    
    iLen = Len(strInput)
    ReDim arByteBlock(iLen - 1)
    
    For i = 1 To iLen
        arByteBlock(i - 1) = Asc(Mid(strInput, i, 1))
    Next i
    
    StringToByteArray = arByteBlock
End Function

Public Function ByteArrayToString(arByteBlock As Variant) As String
    Dim iLB As Long, iUB As Long
    Dim tmp As String
    Dim i As Long
    
    iLB = LBound(arByteBlock)
    iUB = UBound(arByteBlock)
    tmp = ""
    For i = iLB To iUB
        tmp = tmp & Chr(arByteBlock(i))
    Next i
    ByteArrayToString = tmp
End Function

Public Function StringToBinaryString(stInput As String) As String
    Dim iLen As Long
    Dim tmp As String
    Dim i As Long
    
    iLen = Len(Nz(stInput, ""))
    tmp = ""
    
    For i = 0 To iLen - 1
        tmp = tmp & IIf(i = 0, "", ", ") & ByteToBinaryString(Asc(Mid(stInput, i + 1, 1)))
    Next i
    
    StringToBinaryString = tmp
End Function

Public Function StringToLong(stInput As String) As Long
    Dim iLen As Long
    Dim i As Long
    Dim iTmp As Long
    
    iLen = Len(Nz(stInput, ""))
    iTmp = 0
    If iLen > 0 Then
        For i = 0 To iLen - 1
            iTmp = iTmp + Asc(Mid(stInput, i + 1, 1))
        Next i
    End If
    
    StringToLong = iTmp
End Function

Public Function BinaryToByte(arrayBinary As Variant) As Byte
    Dim ub As Integer, i As Byte, r As Byte
    
    ub = UBound(arrayBinary)
    
    If ub > ONE_BYTE - 1 Then
        MsgBox "Subscription " & ub & " is larger than " & ONE_BYTE - 1 & ".", vbInformation, "BinaryToByte"
        Exit Function
    End If
    
    r = 0
    For i = 0 To 7
        r = r + IIf(arrayBinary(i), 2 ^ i, 0)
    Next i
    
    BinaryToByte = r
End Function

Public Function BitAnd(nbrOne As Byte, nbrTwo As Byte) As Byte
    Dim arOne(7) As Boolean, arTwo(7) As Boolean
    Dim i As Byte
    
    For i = 0 To 7
        arOne(i) = False
        arTwo(i) = False
    Next i
    
    ByteToBinary nbrOne, arOne
    ByteToBinary nbrTwo, arTwo
    
    For i = 0 To 7
        arOne(i) = arOne(i) And arTwo(i)
    Next i
    'Debug.Print Asc("A")
    
    BitAnd = BinaryToByte(arOne)
End Function

Public Function BitOr(nbrOne As Byte, nbrTwo As Byte) As Byte
    Dim arOne(7) As Boolean, arTwo(7) As Boolean
    Dim i As Byte
    
    For i = 0 To 7
        arOne(i) = False
        arTwo(i) = False
    Next i
    
    ByteToBinary nbrOne, arOne
    ByteToBinary nbrTwo, arTwo
    
    For i = 0 To 7
        arOne(i) = arOne(i) Or arTwo(i)
    Next i
    'Debug.Print Asc("A")
    
    BitOr = BinaryToByte(arOne)
End Function

'https://code.tutsplus.com/articles/understanding-bitwise-operators--active-11301
Public Function BitShiftLeft(nbrByte As Byte, nbrBits As Byte) As Byte
' Fill right tail with 0
    Dim arByte(7) As Boolean
    Dim i As Byte
    
    nbrBits = Min(nbrBits, 8)
    
    For i = 0 To 7
        arByte(i) = False
    Next i
    
    ByteToBinary nbrByte, arByte
    
    For i = 0 To 7 - nbrBits
        arByte(i) = arByte(i + nbrBits)
    Next i
    
    For i = 7 - nbrBits + 1 To 7
        arByte(i) = False
    Next i
    
    BitShiftLeft = BinaryToByte(arByte)
End Function

Public Function BitShiftRight(nbrByte As Byte, nbrBits As Byte) As Byte
' If left head is One, fill with One. Otherwise, fill with 0.
    Dim arByte(7) As Boolean
    Dim bLeftHead As Boolean
    Dim i As Byte
    
    nbrBits = Min(nbrBits, 8)
    
    For i = 0 To 7
        arByte(i) = False
    Next i
    
    ByteToBinary nbrByte, arByte
    
    bLeftHead = arByte(0)
    
    For i = 7 To nbrBits Step -1
        arByte(i) = arByte(i - nbrBits)
    Next i
    
    For i = 0 To nbrBits - 1
        arByte(i) = bLeftHead
    Next i
    
    BitShiftRight = BinaryToByte(arByte)
End Function

Public Function BitXor(nbrOne As Byte, nbrTwo As Byte) As Byte
    Dim arOne(7) As Boolean, arTwo(7) As Boolean
    Dim i As Byte
    
    For i = 0 To 7
        arOne(i) = False
        arTwo(i) = False
    Next i
    
    ByteToBinary nbrOne, arOne
    ByteToBinary nbrTwo, arTwo
    
    For i = 0 To 7
        arOne(i) = arOne(i) Xor arTwo(i)
    Next i
    'Debug.Print Asc("A")
    
    BitXor = BinaryToByte(arOne)
End Function

Public Sub BitRotateCombined(byteArray As Variant, ByVal shiftBits As Integer, ByVal shiftToLeft As Boolean)
' Takes: an array of byte numbers, number of bits to shift, shift direction.
' Do: Combine all byte numbers into one binary array and shift.
' Returns: Replace passed in byte numbers with shifted numbers.
On Error GoTo Err_BitRotateCombined
    Dim arBoolean() As Variant
    Dim arCom() As Boolean
    Dim iTotalBytes As Integer
    Dim iTotalBits As Integer
    Dim i As Integer, j As Integer
    
    iTotalBytes = UBound(byteArray) + 1
    iTotalBits = iTotalBytes * ONE_BYTE
    
    If iTotalBytes < 1 Or shiftBits = 0 Then
        GoTo Exit_BitRotateCombined
    End If
    
    ReDim arBoolean(iTotalBytes - 1)
    ReDim arCom(iTotalBits - 1)
    
    ' Initialize arBoolean
    For i = 0 To iTotalBytes - 1
        arBoolean(i) = Array(False, False, False, False, False, False, False, False)
    Next i
    
    ' Initialize arCom
    For i = 0 To iTotalBits - 1
        arCom(i) = False
    Next i
    
    shiftBits = shiftBits Mod iTotalBits
    
    For i = 0 To iTotalBytes - 1
        ByteToBinary CByte(byteArray(i)), arBoolean(i)
    Next i
    
    If shiftToLeft Then
        For i = 0 To iTotalBytes - 1
            For j = 0 To ONE_BYTE - 1
                arCom((j + shiftBits + i * ONE_BYTE) Mod iTotalBits) = arBoolean(i)(j)
            Next j
        Next i
    Else
        For i = 0 To iTotalBytes - 1
            For j = 0 To ONE_BYTE - 1
                arCom((IIf((j + i * ONE_BYTE) < shiftBits, iTotalBits, 0) + j + i * ONE_BYTE - shiftBits) Mod iTotalBits) = arBoolean(i)(j)
            Next j
        Next i
    End If
    ' Move result back to arBoolean
    For i = 0 To iTotalBytes - 1
        For j = 0 To ONE_BYTE - 1
            arBoolean(i)(j) = arCom(j + i * ONE_BYTE)
        Next j
    Next i
    ' Convert from binary to byte
    For i = 0 To iTotalBytes - 1
        byteArray(i) = BinaryToByte(arBoolean(i))
    Next i
Exit_BitRotateCombined:
    Exit Sub
Err_BitRotateCombined:
    MsgBox Err.number & ":" & Err.Description, vbCritical, "mMath::BitRotateCombined"
    Resume Exit_BitRotateCombined
End Sub

Public Function HexStringToByte(stHexNumber As String, arReturn() As Byte) As Boolean
On Error GoTo Err_HexStringToByte
    Dim rge As New RegExp
    Dim btLow, btHigh As Byte
    Dim stInput As String
    Dim i As Long, lngLen As Long
    Dim stSource As String
    Dim bSuccess As Boolean
    
    bSuccess = True
    stSource = "StirFry::HexStringToByte"
    stInput = UCase(stHexNumber)
    lngLen = Len(stInput)
    If lngLen < 2 Then
        Err.Raise vbObjectError + 1, stSource, "The length of input string is too short " & _
            "(current length is " & lngLen & ", but minimum 2 is expected)."
    End If
    With rge
        .Multiline = True
        .Global = True
        .Pattern = "[^0-9A-F]+"
        If .Test(stInput) Then
            Err.Raise vbObjectError + 2, stSource, "The input string is not a valid Hexdecimal number"
        End If
    End With
    If (lngLen Mod 2) = 1 Then
        Err.Raise vbObjectError + 3, stSource, "The length of input string must be even number " & _
            "(current length: " & lngLen & ")."
    End If
    
    ReDim arReturn(lngLen / 2 - 1)
    For i = 0 To lngLen / 2 - 1
        btLow = Asc(Mid(stInput, i * 2 + 2, 1))
        btHigh = Asc(Mid(stInput, i * 2 + 1, 1))
        arReturn(i) = btLow + IIf(btLow < 65, -48, -55) + _
            (btHigh + IIf(btHigh < 65, -48, -55)) * 16
    Next i
Exit_HexStringToByte:
    HexStringToByte = bSuccess
    Exit Function
Err_HexStringToByte:
    bSuccess = False
    MsgBox Err.number & ": " & Err.Description, vbCritical, Err.Source
    ReDim arReturn(0)
    Resume Exit_HexStringToByte
End Function

' http://msaccesstipsandtricks.blogspot.com/2014/04/how-to-replace-string-with-regular-expression-in-msacess-vba.html
' https://docs.microsoft.com/en-us/dotnet/standard/base-types/regular-expression-language-quick-reference
' https://stackoverflow.com/questions/5539141/microsoft-office-access-like-vs-regex
Public Function TrimSpaceTab(stInput As String) As String
' Depends on Tools -> References -> "Microsoft VBScript Regular Expressions 5.5"
' \t for tab, \s for space
' ^\t* for leading tab(s), \s*$ for trailing space(s)
On Error GoTo Err_TrimSpaceTab
    Dim rge As New RegExp
    Dim stTmp As String
    
    stTmp = stInput
    With rge
        .Multiline = False
        .IgnoreCase = True
        .Global = False
        .Pattern = "^[\s\t]*"
        stTmp = .Replace(stTmp, "")
        .Pattern = "[\s\t]*$"
        stTmp = .Replace(stTmp, "")
    End With
Exit_TrimSpaceTab:
    Set rge = Nothing
    TrimSpaceTab = stTmp
    Exit Function
Err_TrimSpaceTab:
    Resume Exit_TrimSpaceTab
End Function

Public Function LTrimSpaceTab(stInput As String) As String
On Error GoTo Err_LTrimSpaceTab
    Dim rge As New RegExp
    Dim stTmp As String
    
    stTmp = ""
    With rge
        .Multiline = False
        .IgnoreCase = True
        .Global = False
        .Pattern = "^[\s\t]*"
        stTmp = .Replace(stInput, "")
    End With
Exit_LTrimSpaceTab:
    Set rge = Nothing
    LTrimSpaceTab = stTmp
    Exit Function
Err_LTrimSpaceTab:
    Resume Exit_LTrimSpaceTab
End Function

Public Function RTrimSpaceTab(stInput As String) As String
On Error GoTo Err_RTrimSpaceTab
    Dim rge As New RegExp
    Dim stTmp As String
    
    stTmp = ""
    
    With rge
        .Multiline = False
        .IgnoreCase = True
        .Global = False
        .Pattern = "[\s\t]*$"
        stTmp = .Replace(stInput, "")
    End With
Exit_RTrimSpaceTab:
    Set rge = Nothing
    RTrimSpaceTab = stTmp
    Exit Function
Err_RTrimSpaceTab:
    Resume Exit_RTrimSpaceTab
End Function

Public Function SubStringByPattern(stInput As String, stPattern As String) As String
' Use "(a|ta)\d{4,5}" for Forest File ID/Lincen ID
On Error GoTo Err_SubStringByPattern

    Dim rge As New RegExp
    Dim objTmp As Object
    Dim stTmp As String
    
    stTmp = ""
    
    With rge
        .Multiline = False
        .IgnoreCase = True
        .Global = False
        .Pattern = stPattern
        Set objTmp = .Execute(stInput)
        If objTmp.count > 0 Then
            stTmp = objTmp.Item(0).value
        End If
    End With
    
Exit_SubStringByPattern:
    Set rge = Nothing
    Set objTmp = Nothing
    SubStringByPattern = stTmp
    Exit Function
Err_SubStringByPattern:
    Resume Exit_SubStringByPattern
End Function

Public Function RemoveWhiteSpace(stInput As String) As String
On Error GoTo Err_RemoveWhiteSpace
    Dim rge As New RegExp
    Dim stTmp As String
    
    stTmp = ""
    
    With rge
        .Multiline = False
        .IgnoreCase = True
        .Global = True
        .Pattern = "[\s\t\r\n]"
        stTmp = .Replace(Nz(stInput, ""), "")
    End With
Exit_RemoveWhiteSpace:
    Set rge = Nothing
    RemoveWhiteSpace = stTmp
    Exit Function
Err_RemoveWhiteSpace:
    Resume Exit_RemoveWhiteSpace
End Function
