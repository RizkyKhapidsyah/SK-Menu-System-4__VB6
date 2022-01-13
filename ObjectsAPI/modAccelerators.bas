Attribute VB_Name = "modAccelerators"
Option Explicit
    
''' GUInerd Standard Menu System
''' Version 4.1

''' Objects/API Dll

''' Accelerator Functions


''' ****************************************************************************
''' *** Check the end of the file for copyright and distribution information ***
''' ****************************************************************************

Public Enum AcceleratorValidationResults
    acValid = 1&
    acExists = 2&
    
    acNoExt = &H10&
    acNeedExt = &H20&
    acNoCmd = &H40&
    acNeedCmd = &H80&
    
    acNeedCtrl = &H100&
    acNeedAlt = &H200&
    acNoAlt = &H400&
    
    acDontKnow = &H800&
    acCmdNotAllowed = &H1000&
    
    
    acError = &HFFF0&
    
End Enum

'' Given a keycode, get an item (the code must match precisely and must
'' fit the requirements of the current accelerator processing Style

Public Function KeyCodeToItem(ByVal hWnd As Long, ByVal KeyCode As Long, Optional ByVal Shift As AcceleratorShiftStates, Optional ByVal CommandKey As Long) As MenuItem
    On Error Resume Next
    
    Dim varWnd As Long
    
    Dim varMenu As Object, _
        varItem As MenuItem
        
    For Each varMenu In g_MenuCol
            
        varWnd = varMenu.TopMostParent.hWnd
        
        If (varWnd = hWnd) Then
            For Each varItem In varMenu
            
                With varItem.Accelerator
                
                    If (.KeyCode = KeyCode) And _
                        (.Shift = Shift) And _
                        (.CommandKey = CommandKey) Then
                    
                    
                    
                        If (.CommandKey <> 0&) Then
                            If (CheckPrefix(KeyCode, Shift) = True) Then
                                Set KeyCodeToItem = varItem
                            End If
                        Else
                            If (CheckPrefix(KeyCode, Shift) = False) Then
                                Set KeyCodeToItem = varItem
                            End If
                        End If
                        
                        Exit Function
                    End If
                End With
            
            Next varItem
        End If
        
    Next varMenu
    
End Function

'' Get a menu-item's index from an '&' + selector combination
Public Function SelectorToItem(ByVal lpszSelector As String, ByVal Parent As Object) As Long
    On Error Resume Next
    
    Dim objItem As MenuItem, _
        dwIndex As Long
    
    For Each objItem In Parent
        If (InStr(1, UCase(objItem.Caption), "&" + UCase(lpszSelector)) <> 0&) Then
            SelectorToItem = dwIndex
            Exit Function
        End If
        
        dwIndex = dwIndex + 1
    Next objItem
    
    SelectorToItem = -1&
End Function

'' Convert a Virtual Key Code + Shift combination into a string expression
Public Function AccelCodeToString(ByVal KeyCode As Long, ByVal Shift As AcceleratorShiftStates, Optional ByVal CommandKey As Long) As String

    Dim sAccel As String
    
    If (Shift And acShift) Then
        sAccel = sAccel + "Shift+"
    Else
        If (Shift And acLShift) Then
            sAccel = sAccel + "LShift+"
        End If
        
        If (Shift And acRShift) Then
            sAccel = sAccel + "RShift+"
        End If
    End If
    
    If (Shift And acCtrl) Then
        sAccel = sAccel + "Ctrl+"
    Else
        If (Shift And acLCtrl) Then
            sAccel = sAccel + "LCtrl+"
        End If
        
        If (Shift And acRCtrl) Then
            sAccel = sAccel + "RCtrl+"
        End If
    End If
    
''  System Commands Use Alt
''
    If (Shift And acAlt) Then
        sAccel = sAccel + "Alt+"
    Else
        If (Shift And acLAlt) Then
            sAccel = sAccel + "LAlt+"
        End If
    
        If (Shift And acRAlt) Then
            sAccel = sAccel + "RAlt+"
        End If
    End If
          
    sAccel = sAccel + KeyCodeToToken(KeyCode)
          
    If (CommandKey <> 0&) Then
        sAccel = sAccel + ", " + KeyCodeToToken(CommandKey)
    End If
    
    AccelCodeToString = sAccel

End Function

'' Convert a string expression into a virtual key code + shift combination combination

Public Sub StringToAccelCode(ByVal lpszAccelText As String, KeyCode As Long, Shift As AcceleratorShiftStates, CommandKey As Long)

    Dim sToken As String, _
        uKeyCode As Long, _
        uShift As AcceleratorShiftStates, _
        uCommandKey As Long
    
    Dim sText As String
    
    Dim i As Long
    
    '' Command items follow the sequence: [Shift]+[Key], [Command]
    '' Example: Ctrl+K, W
    
    i = InStr(1, lpszAccelText, ", ")
    If (i <> 0&) Then
        
        sText = Mid(lpszAccelText, 1, i - 1)
        sToken = Mid(lpszAccelText, i + 2)
        uCommandKey = KeyTokenToCode(sToken)
        
    Else
        sText = lpszAccelText
    End If
    
    sToken = StrTok(sText, "+")
                
    Do While (sToken <> "")
        
        Select Case sToken
            
'' Acceptable Shift Codes
            
            Case "LShift", "Left Shift":
                uShift = uShift + acLShift
                uShift = uShift And (-1& Xor acShift)
                
            Case "RShift", "Right Shift":
                uShift = uShift + acRShift
                uShift = uShift And (-1& Xor acShift)
                
            Case "Shift":
                uShift = uShift + acShift
                uShift = uShift And (-1& Xor (acRShift + acLShift))
                                
            Case "LCtrl", "Left Ctrl":
                uShift = uShift + acLCtrl
                uShift = uShift And (-1& Xor acCtrl)
                
            Case "RCtrl", "Right Ctrl":
                uShift = uShift + acRCtrl
                uShift = uShift And (-1& Xor acCtrl)
                
            Case "Ctrl":
                uShift = uShift + acCtrl
                uShift = uShift And (-1& Xor (acRCtrl + acLCtrl))
                                                
            Case "LAlt", "Left Alt":
                uShift = uShift + acLAlt
                uShift = uShift And (-1& Xor acAlt)
                
            Case "RAlt", "Right Alt":
                uShift = uShift + acRAlt
                uShift = uShift And (-1& Xor acAlt)
                
            Case "Alt":
                uShift = uShift + acAlt
                uShift = uShift And (-1& Xor (acRAlt + acLAlt))
                                                
            Case Else
                uKeyCode = KeyTokenToCode(sToken)
        
        End Select
            
        sToken = StrTok(vbNullString, "+")
    Loop
    
    KeyCode = uKeyCode
    Shift = uShift
    CommandKey = uCommandKey

End Sub

Public Function KeyTokenToCode(ByVal sToken As String) As Long
    Dim uKeyCode As Long

'' Acceptable Menu Accelerator Key Codes

    Select Case sToken
    
        Case "F1" To "F9", "F10" To "F24"
            uKeyCode = (val(Mid(sToken, 2)) + 111)
    
        Case "RWin", "Right Window", "RWindow"
            uKeyCode = VK_RWIN
    
        Case "LWin", "Left Window", "LWindow", "Window"
            uKeyCode = VK_LWIN
    
        Case "Apps", "Applications", "Menu"
            uKeyCode = VK_APPS
    
        Case "CapsLock"
            uKeyCode = VK_CAPITAL
    
        Case "SpaceBar", "Space", "Spc"
            uKeyCode = VK_SPACE
    
        Case "Numlock"
            uKeyCode = VK_NUMLOCK
    
        Case "Print", "PrtScrn", "Print Screen"
            uKeyCode = VK_PRINT
    
        Case "Insert", "Ins"
            uKeyCode = VK_INSERT
    
        Case "Help"
            uKeyCode = VK_HELP
        
        Case "Delete", "Del"
            uKeyCode = VK_DELETE
    
        Case "~", "`", "BackQuote", "BackQuotes"
            uKeyCode = VK_BACK_QUOTE
    
        Case Chr(34), "'", "Quote", "Quotes"
            uKeyCode = VK_QUOTE
    
        Case ".", ">", "Period", "Greater-Than"
            uKeyCode = VK_PERIOD
    
        Case ",", "<", "Comma", "Less-Than"
            uKeyCode = VK_COMMA
        
        Case "-", "_", "Dash", "Underscore"
            uKeyCode = VK_DASH
        
        Case "=", "+", "Plus", "Equal"
            uKeyCode = VK_EQUAL
    
        Case "[", "{", "LBrace", "LBracket"
            uKeyCode = VK_LEFT_BRACE
    
        Case "]", "}", "RBrace", "RBracket"
            uKeyCode = VK_RIGHT_BRACE
    
        Case "/", "?", "Backslash", "Question"
            uKeyCode = VK_BACK_SLASH
    
        Case "\", "|", "Slash", "Pipe"
            uKeyCode = VK_SLASH
            
        Case Else
            If (Len(sToken) > 1) Then
                uKeyCode = 0&
            Else
                uKeyCode = Asc(sToken)
            End If
        
            Select Case uKeyCode
                
                Case 48 To 57, 65 To 90
                    uKeyCode = uKeyCode
                                            
                Case Else
                    PostAcceleratorErrorMsg acDontKnow, sToken
                    
            End Select
    
    End Select

    KeyTokenToCode = uKeyCode
End Function

Public Function KeyCodeToToken(ByVal KeyCode As Long) As String
    Dim sAccel As String

    Select Case KeyCode
                    
        Case VK_F1 To VK_F24:
            sAccel = sAccel + "F" & (KeyCode - 111&)
        
        Case VK_RWIN:
            sAccel = sAccel + "RWin"
        
        Case VK_LWIN:
            sAccel = sAccel + "LWin"
        
        Case VK_APPS:
            sAccel = sAccel + "Apps"
        
        Case VK_CAPITAL:
            sAccel = sAccel + "CapsLock"
        
        Case VK_SPACE:
            sAccel = sAccel + "SpaceBar"
        
        Case VK_NUMLOCK:
            sAccel = sAccel + "Numlock"
        
        Case VK_PRINT:
            sAccel = sAccel + "PrtScrn"
        
        Case VK_INSERT:
            sAccel = sAccel + "Insert"
        
        Case VK_HELP:
            sAccel = sAccel + "Help"
        
        Case VK_DELETE:
            sAccel = sAccel + "Delete"
            
        Case VK_BACK_QUOTE
            sAccel = sAccel + "BackQuote"
        
        Case VK_QUOTE
            sAccel = sAccel + "Quote"
        
        Case VK_PERIOD
            sAccel = sAccel + "Period"
        
        Case VK_COMMA
            sAccel = sAccel + "Comma"
        
        Case VK_DASH
            sAccel = sAccel + "Dash"
        
        Case VK_EQUAL
            sAccel = sAccel + "="
        
        Case VK_LEFT_BRACE
            sAccel = sAccel + "["
        
        Case VK_RIGHT_BRACE
            sAccel = sAccel + "]"
                    
        Case VK_BACK_SLASH
            sAccel = sAccel + "\"
            
        Case VK_SLASH
            sAccel = sAccel + "/"
                    
        Case 48 To 57, 65 To 90
            sAccel = sAccel + Chr(KeyCode)
            
    End Select
    
    KeyCodeToToken = sAccel
    
End Function

'' Translate the immediate keyboard shift states into an enum'ed format
Public Function KeyStateImmediate(Optional ByVal LeftRightDifferentiate As Boolean) As AcceleratorShiftStates
    Dim Shift As AcceleratorShiftStates
    
    If LeftRightDifferentiate = True Then
        Shift = 0&
        
        If (GetKeyState(VK_LSHIFT) < 0&) Then
            Shift = acLShift
        End If
        
        If (GetKeyState(VK_RSHIFT) < 0&) Then
            Shift = Shift + acRShift
        End If
        
        If (GetKeyState(VK_LCONTROL) < 0&) Then
            Shift = Shift + acLCtrl
        End If
        
        If (GetKeyState(VK_RCONTROL) < 0&) Then
            Shift = Shift + acRCtrl
        End If

    Else
        Shift = 0&
        
        If (GetKeyState(VK_SHIFT) < 0&) Then
            Shift = acShift
        End If
        
        If (GetKeyState(VK_CONTROL) < 0&) Then
            Shift = Shift + acCtrl
        End If
        
    End If

    KeyStateImmediate = Shift
End Function

Public Function CheckPrefix(ByVal KeyCode As Long, ByVal Shift As AcceleratorShiftStates) As Boolean
    Dim varAccel As Accelerator, _
        fCheck As Boolean
    
    fCheck = False
    
    For Each varAccel In g_MenuCol.PrefixLib
        With varAccel
            If (.KeyCode = KeyCode) And (.Shift = Shift) Then
                fCheck = True
                Exit For
            End If
        End With
    Next varAccel
    
    CheckPrefix = fCheck
    
End Function

Public Function CheckAccelerator(ByVal KeyCode As Long, ByVal Shift As AcceleratorShiftStates, Optional ByVal CommandKey As Long) As Boolean
    Dim varAccel As Accelerator, _
        varMenu As Object, _
        fCheck As Boolean
    
    Dim varAccels As Accelerators
    
    fCheck = False
    
    For Each varMenu In g_MenuCol
        Set varAccels = varMenu.Accelerators
        If Not varAccels Is Nothing Then
            For Each varAccel In varAccels
                With varAccel
                    If (.KeyCode = KeyCode) And (.Shift = Shift) And _
                        (.CommandKey = CommandKey) Then
                        
                        If (CommandKey <> 0&) Then
                            fCheck = CheckPrefix(KeyCode, Shift)
                        Else
                            fCheck = True
                        End If
                        
                        Exit For
                    End If
                End With
            Next varAccel
        
            Set varAccels = Nothing
        End If
    Next varMenu

    Set varAccels = Nothing
    Set varMenu = Nothing
    
    CheckAccelerator = fCheck

End Function

Public Function ValidateAccelerator(ByVal KeyCode As Long, ByVal Shift As AcceleratorShiftStates, ByVal CommandKey As Long, Optional ByVal SysMenu As Boolean, Optional ByVal FromPrefixLib As Boolean) As AcceleratorValidationResults
    
    Dim fExt As Boolean, _
        lResult As AcceleratorValidationResults
    
    Dim fPrefix As Boolean
    
    fExt = ((g_MenuCol.AcceleratorStyle And sbExtendedKeys) = sbExtendedKeys)
    
    lResult = acValid
    
    If (fExt = False) And (Shift And acIsExtended) Then
        lResult = acNoExt
    ElseIf (fExt = True) And Not (Shift And acIsExtended) Then
        lResult = acNeedExt
    End If
    
    fPrefix = CheckPrefix(KeyCode, Shift)
    
    If (FromPrefixLib = True) And (CommandKey <> 0&) Then
        lResult = (lResult + acCmdNotAllowed)
    Else
    
        If (fPrefix = False) And (CommandKey <> 0&) Then
            lResult = (lResult + acNoCmd)
        ElseIf (fPrefix = True) And (CommandKey = 0&) Then
            lResult = (lResult + acNeedCmd)
        End If
    End If
            
    If SysMenu = True Then
        If ((Shift And (acAlt + acLAlt + acRAlt)) = acNone) Then
            lResult = (lResult + acNeedAlt)
        End If
    Else
    
        If ((Shift And (acAlt + acLAlt + acRAlt)) <> acNone) Then
            lResult = (lResult + acNoAlt)
        End If
        
        If ((Shift And (acCtrl + acLCtrl + acRCtrl)) = acNone) Then
            lResult = (lResult + acNeedCtrl)
        End If
    
    End If
    
    If (CheckAccelerator(KeyCode, Shift, CommandKey) = True) Then
        lResult = lResult + acExists
    End If
    
    ValidateAccelerator = lResult
    
End Function


Public Sub PostAcceleratorErrorMsg(ByVal acErrorCode As AcceleratorValidationResults, Optional ByVal acOption As String)

    Dim errMsgA As String, _
        i As Long
        
    Dim errAnd As String
        
    errMsgA = "Sequence "
    
    
    If (acErrorCode And acNoCmd) Then
        If (i >= 1) Then errMsgA = errMsgA + ", "
        errMsgA = errMsgA + "cannot reference a sub-command under an unregistered prefix sequence"
        i = i + 1
    ElseIf (acErrorCode And acNeedCmd) Then
        If (i >= 1) Then errMsgA = errMsgA + ", "
        errMsgA = errMsgA + "references a prefix sequence without a sub-command"
        i = i + 1
    ElseIf (acErrorCode And acCmdNotAllowed) Then
        If (i >= 1) Then errMsgA = errMsgA + ", "
        errMsgA = errMsgA + "would reference a sub-command from the prefix library"
        i = i + 1
    End If
    
    If (acErrorCode And acNoAlt) Then
        If (i >= 1) Then errMsgA = errMsgA + ", "
        errMsgA = errMsgA + "cannot contain the 'Alt' key"
        i = i + 1
    ElseIf (acErrorCode And acNeedAlt) Then
        If (i >= 1) Then errMsgA = errMsgA + ", "
        errMsgA = errMsgA + "must contain the 'Alt' key"
        i = i + 1
    End If
    
    If (acErrorCode And acNeedCtrl) Then
        If (i >= 1) Then errMsgA = errMsgA + ", "
        errMsgA = errMsgA + "must contain the 'Ctrl' key"
        i = i + 1
    End If
    
    If (acErrorCode And acNeedExt) Then
        If (i >= 1) Then errMsgA = errMsgA + ", "
        errMsgA = errMsgA + "must use extended (left and right) keys"
    ElseIf (acErrorCode And acNoExt) Then
        If (i >= 1) Then errMsgA = errMsgA + ", "
        errMsgA = errMsgA + "cannot contain extended (left and right) keys"
    End If
    
    If (i > 1) Then
        errMsgA = "There are several things wrong with this sequence: " + errMsgA
        
        i = InStrRev(errMsgA, ", ")
        
        errAnd = Mid(errMsgA, 1, i + 1)
        errMsgA = Mid(errMsgA, i + 2)
        
        errMsgA = errAnd + "and " + errMsgA
    Else
        errMsgA = "Invalid keyboard accelerator: " + errMsgA
    End If
    
    
    If (acErrorCode = acDontKnow) Then
        errMsgA = "Don't know how to interpret accelerator token: " + Chr(34) + acOption + Chr(34)
    End If
    
    errMsgA = errMsgA + "."
    Err.Raise acErrorCode + &H1600, "StdMenuAPI.modAccelerators", errMsgA
    
End Sub

''' Copyright (C) 2001 Nathan Moschkin

''' ****************** NOT FOR COMMERCIAL USE *****************
''' Inquire if you would like to use this code commercially.
''' Unauthorized recompilation and/or re-release for commercial
''' use is strictly prohibited.
'''
''' please send changes made to code to me at the address, below,
''' if you plan on making those changes publicly available.

''' e-mail questions or comments to nmosch@tampabay.rr.com




