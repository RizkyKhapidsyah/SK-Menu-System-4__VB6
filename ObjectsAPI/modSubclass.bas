Attribute VB_Name = "modSubclass"
''' GUInerd Standard Menu System
''' Version 4.1

''' Objects/API Dll

''' Window Subclassing Management Functions


''' ****************************************************************************
''' *** Check the end of the file for copyright and distribution information ***
''' ****************************************************************************

Option Explicit

Public g_IsSetting As Boolean

Public g_Handles As Collection

Public Function AddHandle(ByVal hWnd As Long, ByVal Proc As Long, Optional ByVal hMenu As Long) As HandleObject
    On Error Resume Next
    
    Dim varNewObj As HandleObject, _
        sKey As String
    
    sKey = "_H" + Hex(hWnd)
    
    If g_Handles Is Nothing Then Set g_Handles = New Collection
    
    Set varNewObj = g_Handles(sKey)
    
    If Not varNewObj Is Nothing Then
        varNewObj.AddReference
        Set varNewObj = Nothing
        
        Exit Function
    End If
    
    Set varNewObj = New HandleObject
    varNewObj.SetHandle hWnd, Proc, hMenu
    
    g_Handles.Add varNewObj, sKey
    Set AddHandle = varNewObj
    Set varNewObj = Nothing
        
End Function

Public Sub RemoveHandle(ByVal hWnd As Long)
    On Error Resume Next
    
    Dim varObj As HandleObject
    Set varObj = g_Handles("_H" + Hex(hWnd))
    
    If Not varObj Is Nothing Then
        varObj.RemoveReference
        
        If (varObj.References = 0&) Then
            g_Handles.Remove "_H" + Hex(hWnd)
        End If
        
        Set varObj = Nothing
    End If
    
    If g_Handles.Count = 0& Then
        Set g_Handles = Nothing
    End If
    
    Set varObj = Nothing
End Sub

Public Function GetOldWndProc(ByVal hWnd As Long) As Long
    On Error Resume Next
    Dim varObj As HandleObject
    
    Set varObj = g_Handles("_H" + Hex(hWnd))
    GetOldWndProc = varObj.OldProc

    Set varObj = Nothing
End Function

Public Function GetPaintDC(ByVal hWnd As Long) As Long
    On Error Resume Next
    Dim varObj As HandleObject
    
    Set varObj = g_Handles("_H" + Hex(hWnd))
    If Not varObj Is Nothing Then
        GetPaintDC = varObj.PaintDC
    End If
    
    Set varObj = Nothing
End Function

'' Main Window Procedure for menus

Public Function MenuWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Static PreCode As Long, _
           PreShift As AcceleratorShiftStates, _
           fPreState As Boolean
    
    On Error Resume Next
    
    Dim Shift As AcceleratorShiftStates, _
        KeyCode As Long
        
    Dim objItem As MenuItem, _
        objBar As Sidebar, _
        objLastHit As MenuItem
    
    Dim dwItemId As Long, _
        dwItemId_Last As Long
    
    Dim OldProc As Long
    
    Dim objMenubar As Menubar, _
        objSubmenu As Submenu, _
        objMenu As Object
            
    Dim objFind As Object, _
        lpInfo As MENUITEMINFO
    
    Select Case uMsg
        
        Case WM_COMMAND, WM_MENURBUTTONUP, WM_SYSCOMMAND
         
            If (uMsg = WM_MENURBUTTONUP) Then
                lpInfo.cbSize = Len(lpInfo)
                lpInfo.fMask = MIIM_ID
                GetMenuItemInfo_API lParam, wParam, True, lpInfo
                
                wParam = lpInfo.wID
            End If
         
            ' Item was selected. ... Where?
            
            ' Find the menu passed by the calling procedure.
            ' To find the menu item id from the WM_COMMAND message,
            ' wParam And'd &Hffff  (low order word of wParam)
            
            If (wParam And &HFFFE0000) Then
                Set objFind = g_MenuCol.FindMenuItem(wParam)
            Else
                Set objFind = g_MenuCol.FindMenuItem(wParam And &HFFFF&)
            End If
            
            If Not objFind Is Nothing Then
                            
                If TypeOf objFind Is MenuItem Then
                    
                    Set objItem = objFind
                    
                    ' The superparent function finds the highest menu
                    ' of the current instance.  Available only as a friend function
                    ' This is a good idea, since the only menus associated with
                    ' window handles are top level instances.
                    
                    objItem.ExecCmd (wParam And &HFFFF&), uMsg
                                        
                    '' Code to activate an MDI child window from the
                    '' window menu of an MDI-frame menu system.
                    
                    If (TypeOf objItem.TopMostParent Is Menubar) And (uMsg <> WM_MENURBUTTONUP) Then
                        If objItem.TopMostParent.Connected = True Then
                            If Not (g_WindowList("_H" + Hex(wParam)) Is Nothing) Then
                            
                                g_IsSetting = True
                                
                                If (g_WinVersion = Windows2000) Or (g_WinVersion = WindowsME) Then
                                    SendMessage objItem.TopMostParent.Connection.hWndClient, WM_MDIACTIVATE, wParam, ByVal 0&
                                Else
                                    SendMessage wParam, WM_ACTIVATE, 1&, ByVal 0&
                                End If
                                
                                g_IsSetting = False
                          
                            End If
                        End If
                    End If
                
                ElseIf ((wParam And &HFFFF&) <> 0&) Then
                
                    Set objBar = objFind
                    objBar.ExecCmd (wParam And &HFFFF&), uMsg
                    
                    Set objBar = Nothing
                    Set objFind = Nothing
                    
                End If
                
            End If
            
        Case WM_INITMENU, WM_INITMENUPOPUP
            
            Set objMenu = g_MenuCol.MenuByHandle(wParam, lParam)
            
            If Not objMenu Is Nothing Then
                SendCommand objMenu, wParam, uMsg
                Set objMenu = Nothing
            End If
        
        Case WM_MENUCHAR
        
            Set objMenu = g_MenuCol("_H" + Hex(lParam))
            
            dwItemId = SelectorToItem(Chr(wParam And &HFFFF&), objMenu)
            MenuWndProc = (MNC_EXECUTE Or dwItemId)
            
            Exit Function
            
        Case WM_TIMER, WM_NCMOUSEMOVE
        
            If (g_WinVersion <> Windows2000) Then
                TestMenubar hWnd, uMsg, wParam, lParam
            End If
                
        Case WM_KEYUP
            
            Shift = KeyStateImmediate()
            KeyCode = wParam
                
            Select Case KeyCode
                Case VK_CONTROL, VK_SHIFT, VK_LCONTROL, VK_LSHIFT, _
                     VK_RCONTROL, VK_RSHIFT
                
                    KeyCode = -1&
            End Select
                
            If (Shift <> 0&) And (KeyCode <> -1&) Then
                
                If (fPreState = True) Then
                    fPreState = False
                    PreCode = 0&
                    PreShift = acNone
                End If
                
                If (CheckPrefix(KeyCode, Shift) = True) Then
                    fPreState = True
                    
                    PreCode = KeyCode
                    PreShift = Shift
                    
                    MenuWndProc = 0&
                Else
                    Set objItem = KeyCodeToItem(hWnd, KeyCode, Shift)
                End If
                
            ElseIf (KeyCode <> -1&) Then
                If (fPreState = True) Or ((KeyCode = VK_ESCAPE) And (Shift = acNone)) Then
                    fPreState = False
                    
                    If (KeyCode <> VK_ESCAPE) Then
                        Set objItem = KeyCodeToItem(hWnd, PreCode, PreShift, KeyCode)
                    End If
                    
                    PreCode = 0&
                    PreShift = acNone
                End If
            End If
            
            If Not objItem Is Nothing Then
                objItem.ExecCmd objItem.ItemId, WM_COMMAND
                Set objItem = Nothing
                Exit Function
            End If
                                          
        Case WM_DRAWITEM
            
            If (g_WinVersion <> Windows2000) Then
                TestMenubar hWnd, uMsg, wParam, lParam
            End If
        
            If (DrawItem(hWnd, lParam) = True) Then
                MenuWndProc = True
                Exit Function
            End If
            
        Case WM_MEASUREITEM
            
            If (MeasureItem(hWnd, lParam) = True) Then
                MenuWndProc = True
                Exit Function
            End If
            
    End Select
    
    Set objFind = Nothing
    Set objItem = Nothing
    Set objBar = Nothing
    
    ' Window procedures enumerated by the class instances by a call to
    ' SetWindowHandle.
    
    OldProc = GetOldWndProc(hWnd)
    
    If (IsBadCodePtr(OldProc) = 0&) Then
        MenuWndProc = CallWindowProc(OldProc, hWnd, uMsg, wParam, lParam)
    Else
        MenuWndProc = DefWindowProc(hWnd, uMsg, wParam, lParam)
    End If
       
End Function

Private Sub TestMenubar(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    On Error Resume Next
    
    Dim lpPoint As POINTAPI, _
        lpDraw As DRAWITEMSTRUCT, _
        lpOrg As POINTAPI, _
        lpRect As RECT

    Dim objItem As MenuItem, _
        objLast As MenuItem
        
    Dim objMenubar As Menubar
    
    Dim hDC As Long, _
        hMenu As Long
    
    Set objMenubar = g_MenuCol.GetMenubar(hWnd)
    
    If objMenubar Is Nothing Then
    
        If (uMsg = WM_TIMER) Then
            Set objMenubar = g_MenuCol("_H" + Hex(wParam))
            If Not objMenubar Is Nothing Then
                KillTimer hWnd, wParam
                Set objMenubar = Nothing
            End If
            
        End If
        
        Exit Sub
    End If
    
    Set objLast = objMenubar.LastOver
    
    hMenu = objMenubar.hMenu
    
    Select Case uMsg
            
        Case WM_TIMER
            
            If (hMenu <> wParam) Then Exit Sub
        
            If Not objLast Is Nothing Then
                
                hDC = GetWindowDC(hWnd)
                
                GetDCOrgEx hDC, lpOrg
                
                objLast.Visual.GetDraw lpDraw
                lpDraw.hDC = hDC
                objLast.Visual.SetDraw lpDraw
                
                With lpDraw.rcItem
                    lpRect.Left = (lpOrg.x + .Left) + 1&
                    lpRect.Top = (lpOrg.y + .Top) + 1
                    
                    lpRect.Right = (lpOrg.x + .Right) - 1&
                    lpRect.Bottom = (lpOrg.y + .Bottom) - 1&
                End With
                
                GetCursorPos lpPoint
                
                If (WithinRect(lpRect, lpPoint.x, lpPoint.y) = False) Then
                    DrawItemIndirect objLast, hWnd, False
                    Set objLast = Nothing
                    
                    KillTimer hWnd, wParam
                End If
                
                ReleaseDC hWnd, hDC
            
            Else
                KillTimer hWnd, wParam
            End If
                    
        Case WM_NCMOUSEMOVE
        
            hDC = GetWindowDC(hWnd)
            
            PointFromLong lParam, lpPoint
            GetDCOrgEx hDC, lpOrg
            
            For Each objItem In objMenubar
                objItem.Visual.GetDraw lpDraw
                
                With lpDraw.rcItem
                    lpRect.Left = (lpOrg.x + .Left) + 1&
                    lpRect.Top = (lpOrg.y + .Top) + 1
                    
                    lpRect.Right = (lpOrg.x + .Right) - 1&
                    lpRect.Bottom = (lpOrg.y + .Bottom) - 1&
                End With
                
                If (WithinRect(lpRect, lpPoint.x, lpPoint.y) = True) Then
                    Exit For
                End If
                
            Next objItem
        
            If (Not objLast Is Nothing) Then
                
                If (objItem Is Nothing) Then
                    KillTimer hWnd, hMenu
                End If
                
                If (Not objLast Is objItem) Then
                    objLast.Visual.GetDraw lpDraw
                    
                    If (lpDraw.itemState And ODS_SELECTED) = 0& Then
                        lpDraw.hDC = hDC
                        objLast.Visual.SetDraw lpDraw
                        DrawItemIndirect objLast, hWnd, 0&
                    
                        Set objLast = Nothing
                    End If
                Else
                    Set objItem = Nothing
                End If
            End If
            
            If (Not objItem Is Nothing) Then
                Set objLast = objItem
                
                objItem.Visual.GetDraw lpDraw
                lpDraw.hDC = hDC
                objItem.Visual.SetDraw lpDraw
                
                DrawItemIndirect objItem, hWnd, ODS_HOTLIGHT
                SetTimer hWnd, hMenu, 50&, 0&
            End If
            
            ReleaseDC hWnd, hDC
            
            Set objItem = Nothing
    
    Case WM_DRAWITEM
            
            Set objLast = Nothing
        
    End Select

    Set objMenubar.LastOver = objLast
    Set objMenubar = Nothing
    Set objLast = Nothing

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



