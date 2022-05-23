VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StirFry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' # File: StirFry
' # Created by: Robert Xie
' # In memory of my dear friend Zhu Qian.(1964 - 2021)
' # Updated: 2022-01-19
' # Requires: mMath module
' # StirFry is an encrypt/decrypt software based on my own algorithm, which is NOT tested for the strength of security. _
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

Const ENCRYPT_METHOD = "SF"     ' Stir Fry
Const ONE_BYTE = 8
Const TWO_BYTES = 16
Const MAX_CHAR = 255
Const MIN_BLOCK_SIZE = 16       ' The minimum encryption block, 128-bits.
Const MAX_BLOCK_SIZE = 512      ' The maximum encryption block, 4096-bits.
Const TO_LEFT = True
Const TO_RIGHT = False
Const TO_FLIP = True
Const TO_FLIP_NOT = False

Dim this_arSeasoning() As Byte
Dim this_block_size As Integer
Dim this_check_sum As Long
Dim this_iPrime As Variant
Dim this_password As String
Dim this_sub_block_size As Integer

Public Function Help(Optional stMethodName As String = "?") As String
    Dim stF As String
    Dim tmp As String
    
    stF = UCase(Nz(stMethodName, "?"))
    stF = IIf(stF = "", "?", stF)
    Select Case stF
    Case "HELP"
        tmp = "In memory of my dear friend Zhu Qian(1964 - 2021)" & vbCrLf & _
            "Usage: Help(Method_Name)" & vbCrLf & _
            "       Help(?) for a list of functions."
    Case "?"
        tmp = "The module contains following functions:" & vbCrLf & _
            "Help" & vbCrLf & _
            "License" & vbCrLf & _
            "Usage: Help(Method_Name)"
    Case "LICENSE"
        tmp = "StirFry is an encrypt/decrypt software based on my own algorithm, which is NOT tested for the strength of security." & vbCrLf & _
            "Copyright (C) 2022 Robert Xie" & vbCrLf & _
            "This program is free software: you can redistribute it and/or modify" & vbCrLf & _
            "it under the terms of the GNU General Public License as published by" & vbCrLf & _
            "the Free Software Foundation, either version 3 of the License, or" & vbCrLf & _
            "(at your option) any later version." & vbCrLf & _
            "This program is distributed in the hope that it will be useful," & vbCrLf & _
            "but WITHOUT ANY WARRANTY; without even the implied warranty of" & vbCrLf & _
            "MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the" & vbCrLf & _
            "GNU General Public License for more details." & vbCrLf & _
            "You should have received a copy of the GNU General Public License" & vbCrLf & _
            "along with this program.  If not, see <https://www.gnu.org/licenses/>."
    Case Else
        tmp = stMethodName & " does not exit in this module."
    End Select
    
    Help = tmp
End Function

Private Sub Class_Initialize()
    this_iPrime = Array(2, 3, 5, 7, 11, 13, 17, 19, 23, 29, 31, 37, 41, 43, 47, _
                53, 59, 61, 67, 71, 73, 79, 83, 89, 97, 101, 103, 107, 109, 113, _
                127, 131, 137, 139, 149, 151, 157, 163, 167, 173, 179, 181, 191, 193, 197, _
                199, 211, 223, 227, 229, 233, 239, 241, 251, 257, 263, 269, 271, 277, 281, _
                283, 293, 307, 311, 313, 317, 331, 337, 347, 349, 353, 359, 367, 373, 379, _
                383, 389, 397, 401, 409, 419, 421, 431, 433, 439, 443, 449, 457, 461, 463, _
                467, 479, 487, 491, 499, 503, 509, 521, 523, 541, 547, 557, 563, 569, 571, _
                577, 587, 593, 599, 601, 607, 613, 617, 619, 631, 641, 643, 647, 653, 659, _
                661, 673, 677, 683, 691, 701, 709, 719, 727, 733, 739, 743, 751, 757, 761, _
                769, 773, 787, 797, 809, 811, 821, 823, 827, 829, 839, 853, 857, 859, 863, _
                877, 881, 883, 887, 907, 911, 919, 929, 937, 941, 947, 953, 967, 971, 977, _
                983, 991, 997, 1009, 1013, 1019, 1021)
    this_block_size = MIN_BLOCK_SIZE
    ReDim this_arSeasoning(this_block_size - 1)
    this_password = ""
End Sub

Public Property Get Block_Size() As Integer
    Block_Size = this_block_size
End Property

Public Property Let Block_Size(intSize As Integer)
    this_block_size = Min(MAX_BLOCK_SIZE, Max(MIN_BLOCK_SIZE, _
        (intSize \ MIN_BLOCK_SIZE + IIf(intSize Mod MIN_BLOCK_SIZE = 0, 0, 1)) * MIN_BLOCK_SIZE))
    
    ReDim this_arSeasoning(this_block_size - 1)
    Prepare_Seasoning
End Property

Public Property Let Password(ByVal stPassword As String)
    Dim lngLen As Long
    
    lngLen = Len(Nz(stPassword, ""))
    If lngLen = 0 Then
        MsgBox "Please input a password.", vbInformation, "StirFry::Password"
        Exit Property
    End If
    this_password = stPassword
    this_sub_block_size = Get_SubBlockSize(stPassword)
    this_check_sum = StringToLong(stPassword)
    Prepare_Seasoning
End Property

Public Function Encrypt(stInput As String, Optional stSalt As String = "") As String
' Use stSalt to alter the shift.
On Error GoTo Err_Encrypt
    Dim lngInputLen As Long
    Dim lngPwdLen As Long
    Dim i As Long
    Dim tmp As String
    Dim iShift As Integer
    
    lngInputLen = Len(Nz(stInput, ""))
    lngPwdLen = Len(Nz(this_password, ""))
    tmp = ENCRYPT_METHOD & this_block_size & "B"
    
    If lngPwdLen = 0 And lngInputLen > 0 Then
        MsgBox "Please set a password before encrypting.", vbCritical, "StirFry::Encrypt"
    ElseIf lngInputLen > 0 Then
        iShift = Get_Shift(this_check_sum + StringToLong(stSalt))
        
        i = 1
        While i < lngInputLen + 1
            tmp = tmp & Encrypt_Block(Mid(stInput, i, this_block_size - 1), iShift)
            i = i + this_block_size - 1
        Wend
    End If
    
Exit_Encrypt:
    Encrypt = tmp
    Exit Function
Err_Encrypt:
    MsgBox Err.number & ": " & Err.Description, vbCritical, "StirFry::Encrypt"
    Resume Exit_Encrypt
End Function

Public Function EncryptWithKey(stInput As String, stHexKey As String, lngCheckSum As Long, Optional stSalt As String = "") As String
On Error GoTo Err_EncryptWithKey
    Dim lngInputLen As Long
    Dim i As Long
    Dim iShift As Integer
    Dim tmp As String
    Dim bValid As Boolean
    
    lngInputLen = Len(Nz(stInput, ""))
    tmp = ENCRYPT_METHOD & this_block_size & "B"
    
    bValid = HexStringToByte(stHexKey, this_arSeasoning)
    If Not bValid And lngInputLen > 0 Then
        MsgBox "Please set a valid encryption key before encrypting.", vbCritical, "StirFry::EncryptWithKey"
    ElseIf lngInputLen > 0 Then
        this_check_sum = lngCheckSum
        iShift = Get_Shift(this_check_sum + StringToLong(stSalt))
        
        i = 1
        While i < lngInputLen + 1
            tmp = tmp & Encrypt_Block(Mid(stInput, i, this_block_size - 1), iShift)
            i = i + this_block_size - 1
        Wend
    End If
    
Exit_EncryptWithKey:
    EncryptWithKey = tmp
    Exit Function
Err_EncryptWithKey:
    MsgBox Err.number & ": " & Err.Description, vbCritical, "StirFry::Encrypt"
    Resume Exit_EncryptWithKey
End Function

Public Function Decrypt(stInput As String, Optional stSalt As String = "") As String
' Use stSalt to alter the shift.
On Error GoTo Err_Decrypt
    Dim stSource As String
    Dim lngInputLen As Long, lngPwdLen As Long
    Dim iBlockSize As Integer
    Dim i As Long
    Dim tmp As String
    Dim stEncrypted As String
    Dim iShift As Integer

    stSource = "StirFry::Decrypt"
    
    lngInputLen = Len(Nz(stInput, ""))
    lngPwdLen = Len(Nz(this_password, ""))
'    tmp = Left(stInput, i - 1)
    tmp = Get_EncryptInfo(stInput)
    i = Len(tmp)
    
    iBlockSize = CInt(Mid(tmp, Len(ENCRYPT_METHOD) + 1, i - Len(ENCRYPT_METHOD) - 1))
    
    If Left(tmp, 2) <> ENCRYPT_METHOD Or iBlockSize <> this_block_size Then
        MsgBox "The message was not encrypted in Stir Fry " & iBlockSize & " Bytes Block.", vbCritical, stSource
    ElseIf lngPwdLen = 0 And lngInputLen > 0 Then
        MsgBox "Please set a password before decrypting.", vbCritical, stSource
    ElseIf lngInputLen > i Then
        stEncrypted = Right(stInput, lngInputLen - i)
        lngInputLen = Len(stEncrypted)
        iShift = Get_Shift(this_check_sum + StringToLong(stSalt))
        
        tmp = ""
        i = 1
        While i < lngInputLen + 1
            tmp = tmp & Decrypt_Block(Mid(stEncrypted, i, this_block_size), iShift)
            i = i + this_block_size
        Wend
    End If
Exit_Decrypt:
    Decrypt = tmp
    Exit Function
Err_Decrypt:
    MsgBox Err.number & ": " & Err.Description, vbCritical, stSource
    Resume Exit_Decrypt
End Function

Public Function DecryptWithKey(stInput As String, stHexKey As String, lngCheckSum As Long, Optional stSalt As String = "") As String
' Use stSalt to alter the shift.
On Error GoTo Err_DecryptWithKey
    Dim stSource As String
    Dim lngInputLen As Long
    Dim iBlockSize As Integer
    Dim i As Long
    Dim iShift As Integer
    Dim tmp As String
    Dim stEncrypted As String
    Dim bValid As Boolean
    
    stSource = "StirFry::DecryptWithKey"
    
    lngInputLen = Len(Nz(stInput, ""))
    tmp = ENCRYPT_METHOD & this_block_size & "B"
    
    bValid = HexStringToByte(stHexKey, this_arSeasoning)
    this_check_sum = lngCheckSum
    iShift = Get_Shift(this_check_sum + StringToLong(stSalt))
    
    If Not bValid And lngInputLen > 0 Then
        MsgBox "Please set a valid decryption key before decrypting.", vbCritical, stSource
    ElseIf lngInputLen > 0 Then
        this_check_sum = lngCheckSum
    End If
    
    tmp = Get_EncryptInfo(stInput)
    i = Len(tmp)
    
    iBlockSize = CInt(Mid(tmp, Len(ENCRYPT_METHOD) + 1, i - Len(ENCRYPT_METHOD) - 1))
    
    If Left(tmp, 2) <> ENCRYPT_METHOD Or iBlockSize <> this_block_size Then
        MsgBox "The message was not encrypted in Stir Fry " & iBlockSize & " Bytes Block.", vbCritical, stSource
    ElseIf lngInputLen > i Then
        If Not bValid Then
            MsgBox "Please set a valid decryption key before decrypting.", vbCritical, stSource
        Else
            stEncrypted = Right(stInput, lngInputLen - i)
            lngInputLen = Len(stEncrypted)
            tmp = ""
            i = 1
            While i < lngInputLen + 1
                tmp = tmp & Decrypt_Block(Mid(stEncrypted, i, this_block_size), iShift)
                i = i + this_block_size
            Wend
        End If
    End If
Exit_DecryptWithKey:
    DecryptWithKey = tmp
    Exit Function
Err_DecryptWithKey:
    MsgBox Err.number & ": " & Err.Description, vbCritical, stSource
    Resume Exit_DecryptWithKey
End Function

Private Function Encrypt_Block(stInput As String, iShift As Integer) As String
On Error GoTo Err_Encrypt_Block
    Dim iLen As Integer
    Dim i As Integer
    Dim iPrimeCount As Integer
    Dim tmp As String
    Dim byteArray() As Byte
    
    iPrimeCount = UBound(this_iPrime) + 1
    iLen = Len(stInput)
    If iLen = 0 Or iLen > this_block_size - 1 Then
        Encrypt_Block = ""
        Exit Function
    End If
    
    tmp = stInput
    
'   Fill empty section with random numbers
    i = iLen
    While i < this_block_size - 1
        tmp = tmp & Chr(Int((MAX_CHAR + 1) * Rnd(0)))
        i = i + 1
    Wend
    tmp = tmp & Chr(iLen)
    
    ReDim byteArray(this_block_size - 1)
    byteArray = StringToByteArray(tmp)
    
    ' First round of applying the password
    DoStirFry byteArray, iShift, TO_LEFT
    
    ' Shuffle use primary numbers for bits.
    ShuffleStringBlock byteArray, UBound(byteArray) + 1, iShift, TO_LEFT
    ShuffleStringBlock byteArray, this_sub_block_size, iShift, TO_LEFT
    
    ' Second round of applying the password
    DoStirFry byteArray, iShift, TO_LEFT, TO_FLIP
Exit_Encrypt_Block:
    Encrypt_Block = ByteArrayToString(byteArray)
    Exit Function
Err_Encrypt_Block:
    MsgBox Err.number & ": " & Err.Description, vbCritical, "StirFry::Encrypt_Block"
    Resume Exit_Encrypt_Block
End Function

Private Function Decrypt_Block(stInput As String, iShift As Integer) As String
On Error GoTo Err_Decrypt_Block
    Dim iLen As Integer
    Dim i As Integer
    Dim iPrimeCount As Integer
    Dim byteArray() As Byte
    
    iPrimeCount = UBound(this_iPrime) + 1
    iLen = Len(stInput)
    If iLen <> this_block_size Then
        Decrypt_Block = ""
        Exit Function
    End If
    
    ReDim byteArray(this_block_size - 1)
    byteArray = StringToByteArray(stInput)
    
    ' Second round of applying the password
    DoStirFry byteArray, iShift, TO_RIGHT, TO_FLIP
    
    ' Shuffle use primary numbers for bits.
    ShuffleStringBlock byteArray, this_sub_block_size, iShift, TO_RIGHT
    ShuffleStringBlock byteArray, UBound(byteArray) + 1, iShift, TO_RIGHT
    
    ' First round of applying the password
    DoStirFry byteArray, iShift, TO_RIGHT
    
Exit_Decrypt_Block:
    Decrypt_Block = Mid(ByteArrayToString(byteArray), 1, byteArray(iLen - 1))
    Exit Function
Err_Decrypt_Block:
    MsgBox Err.number & ": " & Err.Description, vbCritical, "StirFry::Decrypt_Block"
    Resume Exit_Decrypt_Block
End Function

Public Function Get_EncryptInfo(stEncryptSource As String) As String
On Error GoTo Err_Get_EncryptInfo
    Dim tmp As String
    tmp = Nz(stEncryptSource, "")
    tmp = Left(tmp, InStr(tmp, "B"))
Exit_Get_EncryptInfo:
    Get_EncryptInfo = tmp
    Exit Function
Err_Get_EncryptInfo:
    MsgBox Err.number & ": " & Err.Description, vbCritical, "StirFry::Get_EncryptInfo"
    Resume Exit_Get_EncryptInfo
End Function

Public Function Get_EncryptKey() As String
    Dim tmp As String
    Dim stB As String
    Dim i As Integer
    
    tmp = ""
    For i = 0 To this_block_size - 1
        stB = Hex(this_arSeasoning(i))
        tmp = tmp & IIf(Len(stB) = 2, "", "0") & stB
    Next i
    
    Get_EncryptKey = tmp
End Function

Public Property Get Get_CheckSum() As Long
    Get_CheckSum = this_check_sum
End Property

Private Function Get_Shift(lngIndex As Long) As Integer
    Get_Shift = this_iPrime(lngIndex Mod Min((this_block_size - 1) * 8, UBound(this_iPrime)))
End Function

Private Function Get_SubBlockSize(stSeed As String) As Integer
    Get_SubBlockSize = this_iPrime((StringToLong(stSeed) Mod 4) + 1)
End Function

Private Sub Prepare_Seasoning()
' Populate this_arSeasoning
On Error GoTo Err_Prepare_Seasoning
    Dim iPwdLen As Integer, iLen As Integer
    Dim i As Integer, j As Integer
    Dim iShift As Integer
    Dim iCount As Integer
    Dim tmp As String
    Dim arSeasoning() As Byte
    Dim arByte(0) As Byte
    Dim iPrimeCount As Integer
    
    iPrimeCount = UBound(this_iPrime) + 1
    iPwdLen = Len(this_password)
    iLen = Min(iPwdLen, this_block_size)
    iShift = Get_Shift(this_check_sum)
   
    ReDim arSeasoning(this_block_size - 1)
    For i = 0 To iLen - 1
        arSeasoning(i) = Asc(Mid(this_password, i + 1, 1))
    Next i
    
    Select Case iPwdLen
    Case Is < this_block_size
        i = iPwdLen
        iCount = 0
        While i < this_block_size
            For j = 1 To Min(iPwdLen, this_block_size - i, this_block_size)
                arSeasoning(i + j - 1) = Asc(Mid(this_password, j, 1))
                If j Mod 2 = 0 Then
                    BitRotateCombined arSeasoning, this_iPrime(j Mod iPrimeCount), TO_LEFT
                Else
                    ShuffleStringBlock arSeasoning, this_sub_block_size, iShift, TO_LEFT
                End If
            Next j
            iCount = iCount + 1
            i = i + iPwdLen
        Wend
    Case this_block_size
    Case Is > this_block_size
        i = this_block_size
        iCount = 0
        While i < iPwdLen
            ' Spin an over flow character, than xor with a char in coresponding position in the seed.
            For j = 0 To Min(iPwdLen - i, this_block_size - 1)
                arByte(0) = Asc(Mid(this_password, i + j, 1))
                BitRotateCombined arByte, 2 * j + 1, TO_LEFT
                arSeasoning(j) = BitXor(Asc(arSeasoning(j)), arByte(0))
                If j Mod 2 = 0 Then
                    BitRotateCombined arSeasoning, this_iPrime(j Mod iPrimeCount), TO_LEFT
                Else
                    ShuffleStringBlock arSeasoning, 3, iShift, TO_LEFT
                End If
            Next j
            iCount = iCount + 1
            i = i + this_block_size
        Wend
    End Select
    
    ' Spin last character, than xor with the first char in the seed.
    For i = 0 To (this_block_size - 1) / 2
        j = this_block_size - 1 - i
        arByte(0) = arSeasoning(j)
        BitRotateCombined arByte, 2 * i + 1, TO_LEFT
        arSeasoning(i) = BitXor(Asc(arSeasoning(i)), arByte(0))
        If i Mod 2 = 0 Then
            BitRotateCombined arSeasoning, this_iPrime(i Mod iPrimeCount), TO_LEFT
        Else
            ShuffleStringBlock arSeasoning, 3, iShift, TO_LEFT
        End If
    Next i

    For i = 0 To this_block_size - 1
        this_arSeasoning(i) = arSeasoning(i)
    Next
Exit_Prepare_Seasoning:
    Exit Sub
Err_Prepare_Seasoning:
    MsgBox Err.number & ": " & Err.Description, vbCritical, "StirFry::Prepare_Seasoning"
    Resume Exit_Prepare_Seasoning
End Sub

Public Sub ShuffleStringBlock(arByteBlock As Variant, ByVal subBlockSize As Integer, _
                            ByVal nbrBits As Integer, ByVal toLeft As Boolean)
' Takes: a string block; the number of charactors to be grouped when shuffling;
'       rotation bits; rotate direction.
' Do: Take one byte from the head, and one from the tail until the subBlockSize to form a sub block.
'       Rotate the block.
' Return: Replace the original strBlock with the new one.
On Error GoTo Err_ShuffleStringBlock
    Dim iBlockSize As Integer
    Dim arSubBytes() As Byte
    Dim iEffSubBlockSize As Integer
    Dim stOne As String
    Dim i As Integer, j As Integer
    
    If subBlockSize < 1 Then
        Exit Sub
    End If
    
    iBlockSize = UBound(arByteBlock) - LBound(arByteBlock) + 1
    iEffSubBlockSize = Min(iBlockSize, subBlockSize)
    ReDim arSubBytes(iEffSubBlockSize - 1)
    
    i = 0
    While i < iBlockSize
        For j = 0 To iEffSubBlockSize - 1
            If ((i + j) Mod 2) = 0 Then
                arSubBytes(j) = CByte(arByteBlock(Int((i + j) / 2)))
            Else
                arSubBytes(j) = CByte(arByteBlock(iBlockSize - 1 - Int((i + j) / 2)))
            End If
        Next j
        
        BitRotateCombined arSubBytes, nbrBits, toLeft
        
        For j = 0 To iEffSubBlockSize - 1
            If ((i + j) Mod 2) = 0 Then
                arByteBlock(Int((i + j) / 2)) = arSubBytes(j)
            Else
                arByteBlock(iBlockSize - 1 - Int((i + j) / 2)) = arSubBytes(j)
            End If
        Next j
        i = i + iEffSubBlockSize
        If (iBlockSize - i > 0) And (iBlockSize - i < iEffSubBlockSize) Then
            iEffSubBlockSize = iBlockSize - i
            ReDim arSubBytes(iEffSubBlockSize - 1)
        End If
    Wend
Exit_ShuffleStringBlock:
    Exit Sub
Err_ShuffleStringBlock:
    MsgBox Err.number & ": " & Err.Description, vbCritical, "StirFry::ShafflesStringBlock"
    Resume Exit_ShuffleStringBlock
End Sub

Private Sub DoStirFry(ByRef byteCai() As Byte, iShift As Integer, bToLeft As Boolean, Optional bFlip As Boolean = False)
    Dim i As Integer
    Dim iPrimeCount As Integer
    Dim iUB As Integer
    Dim arSeasoning() As Byte
        
    iPrimeCount = UBound(this_iPrime) + 1
    iUB = UBound(this_arSeasoning)
    ReDim arSeasoning(iUB)
    
    If bFlip Then
        For i = 0 To iUB
            arSeasoning(i) = this_arSeasoning(iUB - i)
        Next i
    Else
        For i = 0 To iUB
            arSeasoning(i) = this_arSeasoning(i)
        Next i
    End If
    
    If bToLeft Then
        For i = 0 To iUB
            byteCai(i) = BitXor(byteCai(i), arSeasoning(i))
            If i Mod 2 = 0 Then
                BitRotateCombined byteCai, this_iPrime(i Mod iPrimeCount), TO_LEFT
            Else
                ShuffleStringBlock byteCai, this_sub_block_size, iShift, TO_LEFT
            End If
        Next i
    Else
        For i = iUB To 0 Step -1
            If i Mod 2 = 0 Then
                BitRotateCombined byteCai, this_iPrime(i Mod iPrimeCount), TO_RIGHT
            Else
                ShuffleStringBlock byteCai, this_sub_block_size, iShift, TO_RIGHT
            End If
            byteCai(i) = BitXor(byteCai(i), arSeasoning(i))
        Next i
    End If
End Sub
