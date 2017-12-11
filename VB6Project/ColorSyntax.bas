Attribute VB_Name = "ColorSyntax"
Option Explicit
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long

Const MaxKeyWordsPerLetter = 25
Const MaxAuxDpyLimit = 36

Const ASC_lower_a = 97    ' Print Asc("a")
Const ASC_lower_z = 122   ' Print Asc("z")
Const ASC_upper_a = 65    ' Print Asc("A")
Const ASC_upper_z = 90   ' Print Asc("Z")
Const ASC_zero = 48
Const ASC_nine = 57
Const ASC_dot = 46

' Our lookup table of items
Private LookUp(26, MaxKeyWordsPerLetter) As ColorWords


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: CheckSubRange
' PURPOSE: Checks the lines near where we're editing
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function CheckSubRange(rtb As RichTextBox)
    Dim szTmp As String
    szTmp = rtb.Text
    Dim nOldSelStart As Integer
    nOldSelStart = rtb.SelStart
    CheckRange rtb, PrevCR(szTmp, nOldSelStart), NextCR(szTmp, nOldSelStart)
End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: CheckRange
' PURPOSE: Checks a specific range.
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function CheckRange(rtb As RichTextBox, nStart As Integer, nMax As Integer)

    Dim nOldSelStart, nOldSelLength As Integer
    Dim nPos As Integer, nLen As Integer
    Dim szTmp As String
    Dim Color As ColorWords
    Dim nPercent As Long
    
    With rtb
    
        ' Lock window update
        LockWindowUpdate .hWnd
    
        ' Store current selection
        nOldSelLength = .SelLength
        nOldSelStart = .SelStart
        
        ' Baseline
        If nMax - nStart > 0 Then
            .SelStart = nStart
            .SelLength = nMax - nStart
            .SelColor = vbBlack
        End If

        szTmp = .Text

        ' Find next word
        nPos = NextWord(szTmp, nStart)
        nLen = NextDelimiter(szTmp, nPos + 1) - nPos
        
        While nPos < nMax
        
            Dim strWord As String
            strWord = Mid(szTmp, nPos + 1, nLen)
            
            ' Comment, special case, go to end of line
            If Left(strWord, 1) = "!" Then
                Dim nNextCRPos As Integer
                nNextCRPos = NextCR(szTmp, nPos)
                nLen = nNextCRPos - nPos
                .SelStart = nPos
                .SelLength = nLen
                .SelColor = RGB(0, 128, 0)
                
            Else
                ' Not a comment, see if it's a keyword
                Set Color = IsKeyWord(strWord)
                ' And if it is...
                If Not (Color Is Nothing) Then
                    ' Set color
                    .SelStart = nPos
                    .SelLength = nLen
                    .SelColor = Color.Color
                    
                    ' aux, dpy are string functions
                    ' special cases, they run to the end of the line
                    ' (39 characters max!)
                    If Color.StringType Then
                        nNextCRPos = NextCR(szTmp, nPos)
                        
                        ' Over limit
                        If (nNextCRPos - nPos) > MaxAuxDpyLimit Then
                            nPos = nPos + MaxAuxDpyLimit
                            .SelStart = nPos
                            .SelLength = nNextCRPos - nPos
                            .SelColor = RGB(255, 0, 0)
                        End If
                        nLen = nNextCRPos - nPos
                        
                    End If
                    
                ' TODO: this may not be needed, since the range was set once alreay
                'Else
                '    .SelStart = nPos
                '    .SelLength = nLen
                '    .SelColor = vbBlack
                End If
            End If

            ' Find next word
            nPos = NextWord(szTmp, nPos + nLen)
            nLen = NextDelimiter(szTmp, nPos + 1) - nPos
        Wend
        
        .SelStart = nOldSelStart
        .SelLength = nOldSelLength

        ' Unlock window update
        LockWindowUpdate 0

    End With

End Function



' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: IsKeyWord
' PURPOSE: Check a string to be a keyword
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function IsKeyWord(str As String) As ColorWords

    Dim lstr As String
    lstr = LCase(str)

    Dim n As Integer
    n = GetArrayIndex(str)
    
    ' Assume not
    Set IsKeyWord = Nothing
    
    ' Find the next empty spot
    Dim i As Integer
    For i = 0 To MaxKeyWordsPerLetter
        If LookUp(n, i) Is Nothing Then
            Exit Function
        ElseIf LookUp(n, i).Word = lstr Then
            Set IsKeyWord = LookUp(n, i)
            Exit Function
        End If
    Next

End Function


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: AddWord
' PURPOSE: Adds a word to look up
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Sub AddWord(strWord As String, lColor As Long, Optional bStringType As Boolean = False)
    Dim n As Integer
    n = GetArrayIndex(strWord)
    
    ' Find the next empty spot
    Dim i As Integer
    For i = 0 To MaxKeyWordsPerLetter
        If LookUp(n, i) Is Nothing Then
            Dim oColorWord As New ColorWords
            oColorWord.Color = lColor
            oColorWord.Word = strWord
            oColorWord.StringType = bStringType
            Set LookUp(n, i) = oColorWord
            Exit Sub
        End If
    Next
End Sub


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: LoadSyntax
' PURPOSE:
'   It assumes that the file has a certain structure and
'   performs no error checking. So, be careful editing the
'   files
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub LoadSyntax(file As String)

    ' keyword/command
    AddWord "aux", RGB(0, 0, 128), True
    AddWord "atog", RGB(0, 0, 128)
    AddWord "auto", RGB(0, 0, 128)
    AddWord "and", RGB(0, 0, 128)
    AddWord "bus", RGB(0, 0, 128)
    AddWord "cpl", RGB(0, 0, 128)
    AddWord "dpy", RGB(0, 0, 128), True
    AddWord "dtog", RGB(0, 0, 128)
    AddWord "dec", RGB(0, 0, 128)
    AddWord "declarations", RGB(0, 0, 128)
    AddWord "execute", RGB(0, 0, 128)
    AddWord "goto", RGB(0, 0, 128)
    AddWord "if", RGB(0, 0, 128)
    AddWord "include", RGB(0, 0, 128)
    AddWord "io", RGB(0, 0, 128)
    AddWord "inc", RGB(0, 0, 128)
    AddWord "label", RGB(0, 0, 128)
    AddWord "learn", RGB(0, 0, 128)
    AddWord "pod", RGB(0, 0, 128)
    AddWord "probe", RGB(0, 0, 128)
    AddWord "program", RGB(0, 0, 128)
    AddWord "ram", RGB(0, 0, 128)
    AddWord "ramp", RGB(0, 0, 128)
    AddWord "read", RGB(0, 0, 128)
    AddWord "rom", RGB(0, 0, 128)
    AddWord "run", RGB(0, 0, 128)
    AddWord "stop", RGB(0, 0, 128)
    AddWord "sync", RGB(0, 0, 128)
    AddWord "shl", RGB(0, 0, 128)
    AddWord "shr", RGB(0, 0, 128)
    AddWord "test", RGB(0, 0, 128)
    AddWord "walk", RGB(0, 0, 128)
    AddWord "write", RGB(0, 0, 128)
   
   
    ' variable related colors
    AddWord "assign", RGB(128, 0, 0)
    AddWord "to", RGB(128, 0, 0)
    AddWord "reg0", RGB(128, 0, 0)
    AddWord "reg1", RGB(128, 0, 0)
    AddWord "reg2", RGB(128, 0, 0)
    AddWord "reg3", RGB(128, 0, 0)
    AddWord "reg4", RGB(128, 0, 0)
    AddWord "reg5", RGB(128, 0, 0)
    AddWord "reg6", RGB(128, 0, 0)
    AddWord "reg7", RGB(128, 0, 0)
    AddWord "reg8", RGB(128, 0, 0)
    AddWord "reg9", RGB(128, 0, 0)
    AddWord "reg0", RGB(128, 0, 0)
    AddWord "rega", RGB(128, 0, 0)
    AddWord "regb", RGB(128, 0, 0)
    AddWord "regc", RGB(128, 0, 0)
    AddWord "regd", RGB(128, 0, 0)
    AddWord "rege", RGB(128, 0, 0)
    AddWord "regf", RGB(128, 0, 0)
    
    ' Setup colors
    ' exercise errors
    AddWord "setup", RGB(128, 128, 0)
    AddWord "address", RGB(128, 128, 0)
    AddWord "space", RGB(128, 128, 0)
    AddWord "information", RGB(128, 128, 0)
    AddWord "active", RGB(128, 128, 0)
    AddWord "data", RGB(128, 128, 0)
    AddWord "enable", RGB(128, 128, 0)
    AddWord "error", RGB(128, 128, 0)
    AddWord "errors", RGB(128, 128, 0)
    AddWord "exercise", RGB(128, 128, 0)
    AddWord "force", RGB(128, 128, 0)
    AddWord "interrupt", RGB(128, 128, 0)
    AddWord "line", RGB(128, 128, 0)
    AddWord "trap", RGB(128, 128, 0)
    AddWord "no", RGB(128, 128, 0)
    AddWord "yes", RGB(128, 128, 0)
    
    ' PODs
    AddWord "1802", RGB(0, 128, 128)
    AddWord "1802.pod", RGB(0, 128, 128)
    AddWord "6502", RGB(0, 128, 128)
    AddWord "6502.pod", RGB(0, 128, 128)
    AddWord "6800", RGB(0, 128, 128)
    AddWord "6800.pod", RGB(0, 128, 128)
    AddWord "68000", RGB(0, 128, 128)
    AddWord "68000.pod", RGB(0, 128, 128)
    AddWord "6802", RGB(0, 128, 128)
    AddWord "6802.pod", RGB(0, 128, 128)
    AddWord "6809", RGB(0, 128, 128)
    AddWord "6809.pod", RGB(0, 128, 128)
    AddWord "6809e", RGB(0, 128, 128)
    AddWord "6809e.pod", RGB(0, 128, 128)
    AddWord "8041", RGB(0, 128, 128)
    AddWord "8041.pod", RGB(0, 128, 128)
    AddWord "8048", RGB(0, 128, 128)
    AddWord "8048.pod", RGB(0, 128, 128)
    AddWord "8080", RGB(0, 128, 128)
    AddWord "8080.pod", RGB(0, 128, 128)
    AddWord "8085", RGB(0, 128, 128)
    AddWord "8085.pod", RGB(0, 128, 128)
    AddWord "8086", RGB(0, 128, 128)
    AddWord "8086.pod", RGB(0, 128, 128)
    AddWord "8086mx", RGB(0, 128, 128)
    AddWord "8086mx.pod", RGB(0, 128, 128)
    AddWord "8088", RGB(0, 128, 128)
    AddWord "8088.pod", RGB(0, 128, 128)
    AddWord "8088mx", RGB(0, 128, 128)
    AddWord "8088mx.pod", RGB(0, 128, 128)
    AddWord "9900", RGB(0, 128, 128)
    AddWord "9900.pod", RGB(0, 128, 128)
    AddWord "z80", RGB(0, 128, 128)
    AddWord "z80.pod", RGB(0, 128, 128)
    
   
End Sub



' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: GetArrayIndex
' PURPOSE:
'   This function checks with which letter a keyword starts
'   and returns its index (values 0 to 26)
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function GetArrayIndex(str As String) As Integer

    Dim str1 As String
    str1 = LCase(Left(str, 1))
    
    Dim a As Integer
    a = (Asc(str1) - ASC_lower_a) + 1
    
    If a < 0 Or a > 26 Then
        a = 0
    End If
    
    GetArrayIndex = a

End Function



' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: IsChar
' PURPOSE:
'   Checks if the first character of the string is a non-capital letter
'   Thus, its ascii value lies in the range of 'a'-'z'
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function IsChar(str As String) As Boolean

    Dim strAsc As Integer
    
    strAsc = Asc(str)
    IsChar = strAsc >= ASC_lower_a And strAsc <= ASC_lower_z
    
End Function


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: NextDelimiter
' PURPOSE: Searches a string forward for a delimiter and returns its index
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function NextDelimiter(str As String, start As Integer) As Integer

    Dim tmp As String
    Dim pos As Integer

    If start < 0 Then start = 0

    ' Step trough the string "str" and check every charater
    ' to be a delimiter
    For pos = start + 1 To Len(str)
        tmp = Mid(str, pos, 1)
        If IsDelimiter(tmp) Then
            GoTo done
        End If
    Next pos
   
done:
    NextDelimiter = pos - 1

End Function


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: NextCR
' PURPOSE: Searches a string forward for a CR and returns its index
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function NextCR(str As String, start As Integer) As Integer

    Dim tmp As String
    Dim pos As Integer

    If start < 0 Then start = 0

    ' Step trough the string "str" and check every charater
    ' to be a delimiter
    For pos = start + 1 To Len(str)
        tmp = Mid(str, pos, 1)
        If tmp = vbCr Then
            NextCR = pos - 1
            Exit Function
        End If
    Next pos
   
    NextCR = pos - 1

End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: PrevCR
' PURPOSE: Searches a string backwards for a CR and returns its index
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function PrevCR(str As String, start As Integer) As Integer

    Dim tmp As String
    Dim pos As Integer

    pos = start
    If pos < 0 Then pos = 0

    While pos > 0
        tmp = Mid(str, pos, 2)
        If tmp = vbCrLf Then
            PrevCR = pos + 1
            Exit Function
        End If
        pos = pos - 1
    Wend
    
    PrevCR = 0

End Function

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: IsDelimiter
' PURPOSE: Checks if the first character of the string is a delimiter
' ~@%^&*()-+=\/{}[]:;"'<> ,.?
' Consider chaning this to test a-z, A-Z, and 0-9
' May be faster
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function IsDelimiter(c As String) As Boolean

    Dim nAsc As Integer
    nAsc = Asc(c)
    IsDelimiter = False
    If nAsc >= ASC_lower_a And nAsc <= ASC_lower_z Then Exit Function
    If nAsc >= ASC_upper_a And nAsc <= ASC_upper_z Then Exit Function
    If nAsc >= ASC_zero And nAsc <= ASC_nine Then Exit Function
    If nAsc = Asc(".") Then Exit Function
    ' Special case, this isn't a delimiter, it's treated as a valid character.
    If c = "!" Then Exit Function
    IsDelimiter = True

End Function



' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: NextWord
' PURPOSE: Gets the next word
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Private Function NextWord(str As String, start As Integer) As Integer

    Dim tmp As String
    Dim ret As Boolean
    Dim pos As Integer

    pos = start + 1

    ' Loop until a delimiter is found
    For pos = start + 1 To Len(str)
        tmp = Mid(str, pos, 1)
        If Not IsDelimiter(tmp) Then
            Exit For
        End If
    Next pos
    
    NextWord = pos - 1
    
End Function

