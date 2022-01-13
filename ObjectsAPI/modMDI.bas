Attribute VB_Name = "modMDI"
''' GUInerd Standard Menu System


''' Version 4.1
''' This code is here for reference only
''' MDI support is removed from StdMenuAPI until
''' it can be made stable


''' Objects/API Dll

''' Back-End for the MDI connections and
''' MDI menu collections.


''' ****************************************************************************
''' *** Check the end of the file for copyright and distribution information ***
''' ****************************************************************************


Option Explicit
Option Base 0

Public Type MDICHILDTYPE
    '' handle to window and frame window
    hWnd As Long
    hFrameWnd As Long
    
    '' Handle to menu with window list
    hMenu As Long
    '' Caption to MDI child window as it appears in the window list
    Caption As String
End Type

'' Global WindowList item collection

Public g_WindowList As Collection

'' MDI child collection (for refreshing the WindowList menu)

Public g_MDIWindows() As MDICHILDTYPE

'' Get the client window that holds the menu on the frame window
Public Function GetClientWindow(ByVal hWndFrame As Long) As Long
    
    RefreshChildWindows hWndFrame
    GetClientWindow = g_MDIWindows(0&).hWnd

End Function

'' Get all children created with WS_EX_MDICHILD Style, including
'' the client window (which is always the first window in the list)
Public Function EnumChildWndProc(ByVal hWnd As Long, lParam As Long) As Long
    On Error Resume Next

    Dim G As Long, _
        x As Long
                
    G = GetWindowLong(hWnd, GWL_EXSTYLE)
    x = -1&
    x = UBound(g_MDIWindows)
    
    If (x = -1) Or (G And WS_EX_MDICHILD) Then
        AddMDIChild hWnd, lParam
        EnumChildWndProc = True
    End If
   
End Function

'' Refresh the list for the MDI frame window
Public Sub RefreshChildWindows(ByVal hWndFrame As Long)
    Dim v_Var As Long
    
    Erase g_MDIWindows
    
    v_Var = hWndFrame
    EnumChildWindows hWndFrame, AddressOf EnumChildWndProc, v_Var
    
End Sub

'' Get the frame window for a client/child window
Public Function GetMDIFrame(ByVal hWnd As Long) As Long
    On Error Resume Next

    Dim x As Long
    
    x = -1
    x = UBound(g_MDIWindows)
    
    If x = -1 Then Exit Function
    
    For x = LBound(g_MDIWindows) To UBound(g_MDIWindows)
    
        If g_MDIWindows(x).hWnd = hWnd Then
            GetMDIFrame = g_MDIWindows(x).hFrameWnd
            Exit Function
        End If
            
    Next x
    GetMDIFrame = 0&
    
End Function

'' Update the captions for the windows in the enumeration
'' array

Public Sub UpdateMDICaptions()
    On Error Resume Next
    
    Dim i As Long, _
        j As Long
        
    Dim lpCaption As String
    
    i = -1&
    i = UBound(g_MDIWindows)
    
    If (i = -1&) Then Exit Sub
    
    For j = 0 To i
    
        With g_MDIWindows(j)
            
            lpCaption = String(257, 0&)
            
            GetWindowText .hWnd, lpCaption, 256&
            
            lpCaption = Mid(lpCaption, 1, InStr(1, lpCaption, Chr(0&)) - 1)
    
            .Caption = lpCaption
        End With
    
    Next j
        
End Sub

'' Add an MDI child to the array

Public Function AddMDIChild(ByVal hWnd As Long, ByVal hFrameWnd As Long, Optional ByVal h_Menu As Long)

    On Error Resume Next
    
    Dim x As Long
    Dim lpCaption As String
        
    lpCaption = String(257, 0&)
    GetWindowText hWnd, lpCaption, 256
    lpCaption = Mid(lpCaption, 1, InStr(1, lpCaption, Chr(0&)) - 1)
    
    x = -1&
    x = UBound(g_MDIWindows)
    
    If x = -1 Then
        ReDim g_MDIWindows(0&)
        x = 0&
    Else
        x = x + 1&
        ReDim Preserve g_MDIWindows(0& To x)
    End If
    
    With g_MDIWindows(x)
    
        .hWnd = hWnd
        .hMenu = h_Menu
        .hFrameWnd = hFrameWnd
        
        .Caption = lpCaption
        
    End With
    
End Function

'' Remove an MDI child window from the array

Public Sub RemoveMDIChild(ByVal hWnd As Long)
    On Error Resume Next

    Dim x As Long, y As Boolean
    
    x = -1
    x = UBound(g_MDIWindows)
    
    If x = -1 Then Exit Sub
    
    For x = LBound(g_MDIWindows) To UBound(g_MDIWindows)
    
        If g_MDIWindows(x).hWnd = hWnd Then
        
            y = True
        End If
    
    
        If y = True Then
            
            If (x + 1) <= UBound(g_MDIWindows) Then
                CopyMemory g_MDIWindows(x), g_MDIWindows(x + 1), Len(g_MDIWindows(x))
            Else
                If (x - 1) >= 0& Then
                    ReDim Preserve g_MDIWindows(x - 1)
                Else
                    Erase g_MDIWindows
                End If
                
                Exit For
            End If
        End If
        
    Next x

End Sub

Public Sub RefreshWindowList(hWnd As Long)
    Dim lpSubmenu As Submenu, _
        i As Long, _
        j As Long
        
    If (g_WindowListHandle <> 0&) And (g_IsSetting = True) Then
    
        UpdateMDICaptions
        Set lpSubmenu = g_MenuCol("_H" + Hex(g_WindowListHandle))
                
        If Not lpSubmenu Is Nothing Then
        
            lpSubmenu.ClearWindowList
            
            i = -1&
            i = UBound(g_MDIWindows)
            
            If (i <> -1&) Then
                    
                For j = 1 To i
        
                    lpSubmenu.Add g_MDIWindows(j).Caption, , , , g_MDIWindows(j).hWnd
        
                Next j
        
            End If
        
        End If
        
    End If
    
    UpdateListChecks , hWnd
    
End Sub

Public Sub UpdateListChecks(Optional ByVal v_Activate As Long, Optional ByVal v_hWnd As Long)
    
    Dim v_Item As MenuItem, v_Menu As Submenu
        
    If (v_Activate = 0&) And (v_hWnd <> 0&) Then
        v_Activate = SendMessage(v_hWnd, WM_MDIGETACTIVE, 0&, 0&)
    End If
        
    If (g_WindowListHandle <> 0&) Then
            
        Set v_Menu = g_MenuCol("_H" + Hex(g_WindowListHandle))
            
        If Not v_Menu Is Nothing Then
                
            For Each v_Item In v_Menu.WindowListCol
            
                If (v_Item.Checkmark = False) Then
                    v_Item.Checkmark = True
                End If
                
                If (v_Item.ItemId <> v_Activate) Then
                    v_Item.Checked = False
                Else
                    v_Item.Checked = True
                End If
            Next v_Item
            
        End If
            
    End If
            
    Set v_Item = Nothing
    Set v_Menu = Nothing
            
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




