Attribute VB_Name = "FlukeSignature"
Option Explicit

Public ByteData() As Byte

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' PROPERTY: Signature
' PURPOSE: Calculate the fluke signature based on teh algorithm first
' developed for the Fluke 9010A.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function Signature(strFileNameAndDirectory As String, BytesToAnalyze As Long) As Long

    ' a place to hold the crc data
    Dim crcData(15) As Integer, i As Integer
    
    ' Initialize array
    For i = 0 To 15
        crcData(i) = 0
    Next
    
    ' initialize the pointers, we'll XOR each byte, plus the bytes in the first 4 pointers
    ' into the byte at the last pointer
    Dim pointers(4) As Integer
    pointers(0) = 6
    pointers(1) = 8
    pointers(2) = 11
    pointers(3) = 15
    pointers(4) = 0
    
    On Error GoTo BadFile
    Dim fn As Integer
    Dim FileLength As Long, n As Long
    fn = FreeFile
    Open strFileNameAndDirectory For Binary As fn
    FileLength = LOF(fn)
    ReDim ByteData(0 To FileLength - 1) As Byte
    Get fn, 1, ByteData()
    Close 1
    Dim b As Integer, tmp As Integer
    On Error GoTo 0
    
    For n = 0 To FileLength - 1
        
        tmp = ByteData(n)
    
        For i = 0 To 4
        
            ' get the current pointer
            Dim p As Integer
            p = pointers(i)
            
            ' Either XOr in the data or store the data
            If (i < 4) Then
                tmp = tmp Xor crcData(p)
            Else
                crcData(p) = tmp
            End If
            
            ' move the pointer back one, wrapping around
            p = p - 1
            If (p < 0) Then
                p = 15
            End If
            pointers(i) = p
        
        Next
        
        ' Exit out of the loop if we're asked to stop
        ' at a particular byte. (Note that setting BytesToAnalyze
        ' to zero or -1 will run to the end of the file)
        If (BytesToAnalyze <> 0) And (n = BytesToAnalyze) Then Exit For
        
        
    Next
    
    ' Initialize array
    For i = 0 To 15
        OutputLog.AddOutputLine CStr(i) + ": " + Hex(crcData(i))
    Next

    Dim sig As Long
    sig = 0
    
    ' now for the fun part. processing the crc data into a signature
    For n = 0 To 15
        
       
        tmp = CByte(crcData(n))
        ' for this byte, we need to do it bit by bit
        For i = 0 To 7
        
            Dim bits As Integer
            
            bits = CInt(tmp) Xor (ShiftBits(sig))
        
            ' shift sif left one
            sig = ShiftLeftL(sig, 1)
            
            ' if the 1 bit is set
            sig = sig Or (bits And &H1)
            
            ' shift right 1
            tmp = ShiftRight(tmp, 1)
        Next
    
    Next
    
    ' Finally, our value
    Signature = sig
    Exit Function
    
BadFile:
    ' Bad data.
    Signature = &HFFFF
    ReDim ByteData(0) As Byte
    ByteData(0) = &HFF
    
End Function


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: ShiftBits
' PURPOSE: The CRC polynomial function for fluke
'   C#: return (UInt16)((sig >> 6) ^ (sig >> 8) ^ (sig >> 11) ^ (sig >> 15));
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function ShiftBits(sig As Long) As Long

    Dim lSig As Long
    lSig = sig
    
    Dim s6 As Long, s8 As Long, s11 As Long, s15 As Long
    s6 = ShiftRight(lSig, 6)
    s8 = ShiftRight(lSig, 8)
    s11 = ShiftRight(lSig, 11)
    s15 = ShiftRight(lSig, 15)
    
    ShiftBits = s6 Xor s8 Xor s11 Xor s15
    ShiftBits = ShiftBits And &HFF

    
End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: ShiftLeft
' PURPOSE: <<
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Static Function ShiftLeft(ByVal Value As Long, ByVal ShiftCount As Long) As Long
    ShiftLeft = Value
    While ShiftCount > 0
        ShiftLeft = ShiftLeft And &HFF
        ShiftLeft = ShiftLeft * 2
        ShiftCount = ShiftCount - 1
    Wend
    
    ShiftLeft = ShiftLeft And &HFF
End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: ShiftLeftL1
' PURPOSE: Shifts up (multiplies by 2) for a long type (which could
'   overflow)
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Static Function ShiftLeftL(ByVal Value As Long, ByVal ShiftCount As Long) As Long
    
    ShiftLeftL = Value
    While ShiftCount > 0
        ' Make sure we don't overflow. The high order bit will be lost,
        ' so mask it off
        ShiftLeftL = ShiftLeftL And &H7FFF
        ShiftLeftL = ShiftLeftL * 2
        ShiftCount = ShiftCount - 1
    Wend
    
    ShiftLeftL = ShiftLeftL And &HFFFF
End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: ShiftRight
' PURPOSE: >>
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Static Function ShiftRight(ByVal Value As Long, ByVal ShiftCount As Long) As Long

    ShiftRight = Value
    While ShiftCount > 0
        ' Rounding may occur (because VB sucks) so mask off the
        ' lowest order bit. Doesn't affect bit shift
        ShiftRight = ShiftRight And &HFFFE
        ShiftRight = ShiftRight / 2
        ShiftCount = ShiftCount - 1
    Wend
    
    ShiftRight = ShiftRight And &HFF

End Function

