Attribute VB_Name = "modTextMan"
''' GUInerd Standard Menu System
''' Version 4.1

'''

''' Common Text Manipulation Routines


''' ****************************************************************************
''' *** Check the end of the file for copyright and distribution information ***
''' ****************************************************************************

Option Explicit

Public Const DefaultSeparators = " _-"
Public Const WrapLimit_Minimum = 20&

Public Function RemoveNulls(ByVal szText As String) As String
    Dim i As Long, _
        j As Long
        
    Dim c As String, _
        d As String
    
    i = InStr(1, szText, Chr(0))
    
    If (i = 0&) Then
        RemoveNulls = szText
        Exit Function
    End If
    
    i = Len(szText)
    
    For j = 1 To i
        c = Mid(szText, j, 1)
        If (c <> Chr(0)) And (c <> ChrW(0)) Then
            d = d + c
        End If
    Next j
    
    RemoveNulls = d
    
End Function

'' Determine whether character in buffer is a separator character or not.
Public Function IsSepChar(ByVal sChr As String, Optional ByVal sSeparators As String = DefaultSeparators) As Boolean
    Dim i As Long, _
        j As Long
        
    i = Len(sChr)
    If (i <> 1) Or (sSeparators = "") Then Exit Function
    i = Len(sSeparators)
    
    For j = 1 To i
        If (sChr = Mid(sSeparators, j, 1)) Then
            IsSepChar = True
            Exit Function
        End If
    
    Next j
        
    IsSepChar = False
    
End Function

'' Find the next word/separator(or crlf) pair
Public Function NextWordSep(ByVal vData As String, ByVal uStartPos As Long, lpszText As String, lpszSeparator As String, Optional ByVal szPattern As String, Optional ByVal sCrlf As String = vbCrLf) As Long
    Dim i As Long, _
        j As Long
    
    Dim c As Long, _
        z As Long
    
    Dim uFind As String, _
        uTxt As String
    
    Dim s1 As String, _
        p As Long
    
    If (szPattern = "") Then
        szPattern = DefaultSeparators
    End If
    
    i = uStartPos
    If i = 0& Then i = 1&
    
    j = Len(szPattern)
        
    For c = 1 To j
        s1 = Mid(szPattern, c, 1)
        z = InStr(i, vData, s1)
        
        If (z <> 0&) Then
            If (z < p) Or (p = 0&) Then
                p = z
                uFind = s1
            End If
        End If
    Next c
    
    If (sCrlf <> "") Then
        z = InStr(i, vData, sCrlf)
        
        If (z <> 0&) Then
            If (z < p) Or (p = 0&) Then
                p = z + (Len(sCrlf) - 1)
                uFind = sCrlf
            End If
        End If
    End If
    
    If (p <> 0&) Then
        If (p <> i) Then
            uTxt = Mid(vData, i, p - i)
        Else
            uTxt = ""
        End If
        
        lpszText = uTxt
        lpszSeparator = uFind
        
        If ((p + 1) <= Len(vData)) Then
            NextWordSep = p + 1
        Else
            NextWordSep = -1&
        End If
        
        Exit Function
    End If
    
    lpszText = Mid(vData, i)
    lpszSeparator = ""
    NextWordSep = -1&
    
End Function

'' Parse out all words in a text buffer using the default separators and new-line sequence

Public Function Words(ByVal vData As String) As String()
    Dim i As Long, _
        j As Long
    
    Dim vOut() As String, _
        v1 As String, _
        v2 As String
    
    i = 1&
    Do While i <> -1&
        i = NextWordSep(vData, i, v1, v2)
        
        If v1 <> "" Then
            ReDim Preserve vOut(0& To j)
            vOut(j) = RemoveNulls(v1)
            j = j + 1
        End If
        
        If v2 <> "" Then
            ReDim Preserve vOut(0& To j)
            vOut(j) = RemoveNulls(v2)
            j = j + 1
        End If
    Loop
    
    Words = vOut
    
End Function

'' Wrap Text and Parse Lines (w/ Cr+Lf)

Public Function WrapLines(ByVal hDC As Long, cbLimit As Long, ByVal szText As String, _
                          Optional ByVal sSeparators As String = DefaultSeparators, Optional ByVal sCrlf As String = vbCrLf, _
                          Optional ByVal fHasAccel As Boolean) As String()

    Dim sWords() As String, _
        i As Long, _
        j As Long, _
        lWords As Long, _
        lHeight As Long
    
    Dim sCurrent As String, _
        sPrev As String, _
        sLines() As String, _
        nLines As Long
    
    Dim lpSize As SIZEAPI, _
        lpWordSize As SIZEAPI, _
        bNoInc As Boolean
    
    Dim lWidth As Long
    
    On Error Resume Next
    
    If szText = "" Then Exit Function
    sWords = Words(szText)
    
    i = UBound(sWords)
    
    sCurrent = sWords(0&)
    sPrev = sCurrent
    
    ReDim sLines(0&)
    lWords = 1
    
    Do While sCurrent <> ""
    
        MeasureText_API hDC, sWords(j), lpWordSize, fHasAccel
        MeasureText_API hDC, sCurrent, lpSize, fHasAccel
                
        If (lWords = 1) And (cbLimit >= WrapLimit_Minimum) Then
            If lpSize.cx > cbLimit Then
                cbLimit = lpSize.cx
            End If
        
            If lpWordSize.cx > cbLimit Then
                cbLimit = lpWordSize.cx
            End If
        End If
        
        If (lpSize.cx > lWidth) Then lWidth = lpSize.cx
        
        If ((cbLimit >= WrapLimit_Minimum) And (lpSize.cx > cbLimit)) Or ((sCrlf <> "") And (sWords(j) = sCrlf)) Then
            sLines(nLines) = sPrev
            
            nLines = nLines + 1
            
            ReDim Preserve sLines(0& To nLines)
            
            lWords = 1
            
            If (IsSepChar(sWords(j), sSeparators) = True) Or ((sCrlf <> "") And (sWords(j) = sCrlf)) Then
                sLines(nLines - 1) = sLines(nLines - 1) + sWords(j)
                
                If (j < i) Then
                    sPrev = sWords(j + 1)
                Else
                    nLines = nLines + 1
                    ReDim Preserve sLines(0& To nLines)
                    sPrev = ""
                    bNoInc = False
                End If
                
                j = j + 1
            Else
                sPrev = sWords(j)
            End If
            
            sCurrent = sPrev
            bNoInc = True
        End If
        
        If (j < i) And (bNoInc = False) Then
            sPrev = sCurrent
            
            j = j + 1
            sCurrent = sCurrent + sWords(j)
            lWords = lWords + 1
        Else
            If bNoInc = True Then
                bNoInc = False
            ElseIf (j >= i) Then
                Exit Do
            End If
        End If
    Loop

    sLines(nLines) = sCurrent
    If sLines(nLines) = "" Then
        ReDim Preserve sLines(0& To (nLines - 1))
    Else
        nLines = nLines + 1
    End If
    
    If (cbLimit < WrapLimit_Minimum) Then
        cbLimit = lWidth
    End If
    
    WrapLines = sLines
    
End Function

'' Parse a string and optionally measure/retrieve accelerator info (for use with menus)

Public Function Parse_String(ByVal hDC As Long, ByVal lpszText As String, Optional lpdwLines As Long, Optional lpdwLimit As Long, Optional lpdwAccel As Long, _
                             Optional ByVal lpSize As Long, Optional ByVal sCrlf As String = vbCrLf) As String()
    
    Dim x As String, _
        y As String
        
    Dim wrapStr() As String, _
        outStr() As String
    
    Dim i As Long, _
        j As Long
        
    Dim a As Long, _
        b As Long
        
    Dim nCount As Long, _
        cbText As SIZEAPI
        
    Dim dwLimit As Long, _
        n As Long
                
    Dim fHasAccel As Boolean
                
    On Error Resume Next
    
    Erase outStr
    
    If (IsMissing(lpdwAccel) = False) Then
        fHasAccel = True
    End If
    
    If (IsMissing(lpdwLimit) = False) Then
        dwLimit = lpdwLimit
    End If
    
    outStr = WrapLines(hDC, dwLimit, lpszText, , , fHasAccel)

    b = -1&
    b = UBound(outStr)
    
    nCount = b + 1
        
    If (nCount > 0&) Then
        If (lpSize <> 0&) Then
            MeasureLines_API hDC, outStr, cbText, fHasAccel
            CopyMemory ByVal lpSize, cbText, Len(cbText)
        End If
        
        If IsMissing(lpdwLimit) = False Then
            lpdwLimit = dwLimit
        End If
        
        If IsMissing(lpdwLines) = False Then
            lpdwLines = nCount
        End If
        
        If IsMissing(lpdwAccel) = False Then
            For i = 0& To b
                n = InStr(1, outStr(i), "&")
                If (n <> 0&) Then
                    lpdwAccel = i
                    Exit For
                End If
            Next i
        End If
            
        Parse_String = outStr
        outStr = Empty
    End If
    
End Function

'' re-assemble parsed-out lines

Public Function Reassemble(lpLines() As String) As String
    Dim s1 As String, _
        i As Long, _
        j As Long
            
    On Error Resume Next
    
    i = -1&
    i = UBound(lpLines)
    
    For j = 0& To i
        s1 = s1 + lpLines(j)
    Next j
    
    Reassemble = s1
    
End Function


'' String Token Function (works like in C/C++)
Public Function StrTok(ByVal v_Str As String, ByVal vTok As String) As String
    Static vOldStr As String
    Dim i As Long, l As Long, c As String, d As String
        
    If Not v_Str = vbNullString Then
        vOldStr = v_Str
    End If
    
    If vOldStr = vbNullString Then Exit Function
    
    If Mid(vOldStr, 1, Len(vTok)) = vTok Then
        vOldStr = Mid(vOldStr, Len(vTok) + 1)
        StrTok = ""
        Exit Function
    End If
        
    For i = 1 To Len(vOldStr)
        c = Mid(vOldStr, i, Len(vTok))
        If c = vTok Then
            StrTok = d
            vOldStr = Mid(vOldStr, i + Len(vTok))
            Exit Function
        Else
            c = Mid(vOldStr, i, 1)
            d = d + c
        End If
    Next i
    
    StrTok = d
    vOldStr = vbNullString
End Function

Public Function GetPlainCaption(ByVal vData As String) As String

    Dim i As Long, _
        j As Long, _
        s As String, _
        t As String

    j = Len(vData)
    
    For i = 1 To j
        If (i < (j - 1)) Then
            s = Mid(vData, i, 3)
            If (s = (vbCr + vbCrLf)) Then
                i = i + 3
            End If
        End If
        
        If (i < j) Then
            s = Mid(vData, i, 2)
            
            If Mid(s, 1, 1) = "&" Then
                t = t + Mid(s, 2)
                i = i + 1
            ElseIf s <> vbCrLf Then
                t = t + Mid(s, 1, 1)
            Else
                i = i + 1
            End If
        Else
            t = t + Mid(vData, i)
        End If
    Next i
    
    GetPlainCaption = t
            
End Function

Public Function GetAccelPos(ByVal vData As String) As Long

    Dim i As Long, s As String, t As String
    
    For i = 1 To Len(vData)
        s = Mid(vData, i, 1)
        If s = "&" Then
            If Len(vData) > i Then
                If Mid(vData, i + 1, 1) <> "&" Then
                    GetAccelPos = Len(t) + 1
                End If
            End If
        Else
            t = t + s
        End If
    Next i
    
End Function

Public Function ExtractString(ByVal lpsz As Long, Optional fIsUnicode As Boolean) As String
    Dim xUni As Boolean

    Dim lpStrNew As String, _
        a As Long, _
        b As Long

    Dim tIn As Byte, _
        wIn As Integer

    If IsMissing(fIsUnicode) Then
        If (g_WinVersion = Windows2000) Or (g_WinVersion = WindowsNT) Then
            xUni = True
        Else
            xUni = False
        End If
    Else
        xUni = fIsUnicode
    End If
    
    
    a = lpsz
    
    If (xUni = True) Then
        CopyMemory wIn, ByVal a, 2&
        b = wIn
    Else
        CopyMemory tIn, ByVal a, 1&
        b = tIn
    End If
    
    Do While b <> 0&
        If xUni = True Then
            lpStrNew = lpStrNew + ChrW(b)
            a = a + 2
        Else
            lpStrNew = lpStrNew + Chr(b)
            a = a + 1
        End If
    
        If (xUni = True) Then
            CopyMemory wIn, ByVal a, 2&
            b = wIn
        Else
            CopyMemory tIn, ByVal a, 1&
            b = tIn
        End If
            
    Loop
    
    ExtractString = lpStrNew
End Function


''' Copyright (C) 2001 Nathan Moschkin

''' ****************** NOT FOR COMMERCIAL USE *****************
''' Inquire if you would like to use this code commercially.
''' Unauthorized recompilation and/or re-release for commercial
''' use is strictly prohibited.
'''
''' please send changes made to code to me at the address, below,
''' if you plan on making those changes publicly available.

''' e-mail questions or comments to nmosch@tampabay.rr.com


