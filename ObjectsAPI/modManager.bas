Attribute VB_Name = "modManager"
''' GUInerd Standard Menu System
''' Version 4.1

''' Objects/API Dll

''' General Menu Management and Sub Main()


''' ****************************************************************************
''' *** Check the end of the file for copyright and distribution information ***
''' ****************************************************************************

Option Explicit

'' For information on MENUKEY, see the ChangeActiveMenuSet() subroutine, below.

Private Type MENUKEY
    Key As String
    hWnd As Long
End Type


Public Const CopyableMask = (MFT_STRING Or MFT_OWNERDRAW Or MFT_SEPARATOR Or MFT_MENUBARBREAK Or MFT_MENUBREAK Or MFT_RIGHTJUSTIFY)

Public g_CtrlIdNext As Long

'' ItemInfo size (depends on the value of g_WinVersion)
Public g_InfoSize As Long

'' Item identifier to the active windowlist menu on an MDI menu.
Public g_WindowListHandle As Long

'' Global default menu font
Public g_SysMenuFont As StdFont

'' Global menu collection
Public g_MenuCol As Menus

'' Global Windows Version info
Public g_WinVersion As WindowsVersionConstants

'' For scaling images, etc.
Public ScaleTool As ScaleNumeric


'' For each thread of the DLL, the Main() sub is invoked to initialize
'' global data.

Sub Main()
    
    GetWinVersion
    
    '' The ScaleNumeric object is private.  The object's
    '' internal reference will not affect proper shut-down.
    Set ScaleTool = New ScaleNumeric
    
    '' likewise, references to VB components do not affect shut-down.
    Set g_WindowList = New Collection
    Set g_Handles = New Collection
    
    ' Create a new global default font object
    Set g_SysMenuFont = GetSysMenuFont
    
    If g_SysMenuFont Is Nothing Then
        Set g_SysMenuFont = New StdFont
        g_SysMenuFont.Size = 8#
        g_SysMenuFont.Name = "MS Sans Serif"
    End If
        
    g_CtrlIdNext = &H1248&
    
    Set g_MenuCol = New Menus
    
End Sub

'' Methods to get and set item information
'' with Unicode support

Public Function GetMenuItemInfo_API(ByVal hMenu As Long, ByVal ItemId As Long, ByVal fBool As Boolean, lpInfo As MENUITEMINFO, Optional ItemCaption As String) As Long
    Dim varByte() As Byte, _
        varStr As String
        
    Dim cch As Long
    
    On Error Resume Next
    
    If (Unicode_Checked = False) Then CheckUnicode
    
    lpInfo.cbSize = g_InfoSize
    
    If (lpInfo.fMask And MIIM_TYPE) Then
        cch = 512&
        ReDim varByte(0 To cch)
        lpInfo.dwTypeData = VarPtr(varByte(0))
        lpInfo.cch = cch
    End If
    
    If (Using_Unicode = True) Then
        GetMenuItemInfoW hMenu, ItemId, fBool, lpInfo
        
        If (lpInfo.fType = MFT_STRING) And (lpInfo.fMask And MIIM_TYPE) And (IsMissing(ItemCaption) = False) Then
            cch = lpInfo.cch
            varStr = String(cch, 0)
            
            CopyMemory ByVal StrPtr(varStr), ByVal (lpInfo.dwTypeData), (cch * 2&)
            ItemCaption = varStr
            
        End If
    Else
        GetMenuItemInfo hMenu, ItemId, fBool, lpInfo
    
        If (lpInfo.fType = MFT_STRING) And (lpInfo.fMask And MIIM_TYPE) And (IsMissing(ItemCaption) = False) Then
            cch = lpInfo.cch
            ReDim Preserve varByte(0 To (cch - 1))
                                                
            CopyMemory ByVal VarPtr(varByte(0)), ByVal (lpInfo.dwTypeData), cch
            varStr = StrConv(varByte, vbUnicode)
            
            ItemCaption = varStr
        End If
    End If

    If (lpInfo.fMask And MIIM_TYPE) Then
        Erase varByte
    End If
    
End Function

Public Function SetMenuItemInfo_API(ByVal hMenu As Long, ByVal ItemId As Long, ByVal fBool As Boolean, lpInfo As MENUITEMINFO, Optional ByVal ItemCaption As String) As Long
    Dim varByte() As Byte, _
        varStr As String
        
    Dim cch As Long
    
    On Error Resume Next
    
    If (Unicode_Checked = False) Then CheckUnicode
    
    lpInfo.cbSize = g_InfoSize
    
    If (lpInfo.fType = MFT_STRING) And (lpInfo.fMask And MIIM_TYPE) And (ItemCaption <> "") Then
    
        If (Using_Unicode = True) Then
            varStr = ItemCaption
        Else
            varStr = StrConv(ItemCaption, vbFromUnicode)
        End If
    
        cch = LenB(varStr)
        ReDim varByte(0 To cch)
    
        CopyMemory varByte(0), ByVal StrPtr(varStr), cch
        
        lpInfo.dwTypeData = VarPtr(varByte(0))
        lpInfo.cch = cch
        
    End If

    If (Using_Unicode = True) Then
        SetMenuItemInfoW hMenu, ItemId, fBool, lpInfo
    Else
        SetMenuItemInfo hMenu, ItemId, fBool, lpInfo
    End If

    Erase varByte

End Function

Public Sub RefreshItem(ByVal ItemId As Long, Optional ByVal Reindex As Boolean, Optional ByVal Item As Object)
    Dim varItem As Object, _
        varObj As Object, _
        i As Long, _
        j As Long
        
    Dim c As Long, _
        hTop As Long, _
        cbFlags As Long
        
    Dim hMenu As Long, _
        lpInfo As MENUITEMINFO
        
    Dim iItem As MenuItem, _
        sItem As Sidebar
        
    On Error Resume Next
    
    Set varItem = g_MenuCol.FindMenuItem(ItemId)
    
    If varItem Is Nothing Then
        Set varItem = Item
    End If
    
    hMenu = varItem.Parent.hMenu
    
    If varItem Is Nothing Then Exit Sub
        
    cbFlags = (MF_OWNERDRAW Or MF_BYPOSITION)
        
    If TypeOf varItem Is Sidebar Then
        
        Set sItem = varItem
        sItem.Freeze
        
        RemoveMenu hMenu, ItemId, MF_BYCOMMAND
        RemoveMenu hMenu, sItem.BreakId, MF_BYCOMMAND
        
        If (sItem.Visible = True) Then
            
            If (sItem.Position = posRight) Then
                AppendMenu hMenu, (cbFlags Or MF_MENUBREAK Or MF_SEPARATOR), sItem.BreakId, 0&
                
                If (sItem.Enabled = True) Then
                    AppendMenu hMenu, cbFlags, ItemId, 0&
                Else
                    AppendMenu hMenu, cbFlags Or MF_SEPARATOR, ItemId, 0&
                End If
                
            Else
                If (sItem.Enabled = True) Then
                    InsertMenu hMenu, 0&, cbFlags, ItemId, 0&
                Else
                    InsertMenu hMenu, 0&, cbFlags Or MF_SEPARATOR, ItemId, 0&
                End If
                
                InsertMenu hMenu, 1&, (cbFlags Or MF_MENUBREAK Or MF_SEPARATOR), sItem.BreakId, 0&
            End If
            
            sItem.Unfreeze
            sItem.SetCtrlState (True)
        End If
                            
        Set sItem = Nothing
        Set varItem = Nothing
        
        Exit Sub
    
    End If
    
    Set iItem = varItem
    iItem.Freeze
    
    RemoveMenu hMenu, ItemId, MF_BYCOMMAND
    
    If (iItem.Visible = True) Then
        
        If Not TypeOf iItem.Parent Is Menubar Then
            Set varObj = iItem.Parent.Sidebar
        End If
        
        If Not varObj Is Nothing Then
            If (varObj.Visible = False) Or (varObj.Position = posRight) Then
                InsertMenu hMenu, (iItem.Index), cbFlags, ItemId, 0&
            Else
                InsertMenu hMenu, (iItem.Index + 2&), cbFlags, ItemId, 0&
            End If
        Else
            InsertMenu hMenu, iItem.Index, cbFlags, ItemId, 0&
        End If
            
        iItem.Unfreeze
        iItem.SetCtrlState (True)
    End If
    
    If (Reindex = True) Then ReindexMenu iItem.Parent
                
    Set iItem = Nothing
    Set varItem = Nothing
    Set varObj = Nothing
    
End Sub

Public Sub ReindexMenu(Menu As Object)
    Dim i As Long, _
        j As Long, _
        varItem As MenuItem
    
    Dim x As Long
        
    j = Menu.Count
    
    For i = 1 To j
        Set varItem = Menu.Item(i)
        
        If (i > 1) Then
            Set varItem.PrevItem = Menu.Item(i - 1)
        End If
        
        If (i < j) Then
            Set varItem.NextItem = Menu.Item(i + 1)
        End If
        
        varItem.Index = x
        x = x + 1
        
        Set varItem = Nothing
    Next i
        
    UpdateTables Menu
            
End Sub

Public Sub UpdateTables(ByVal Menu As Object)
    Dim objAccels As Accelerators, _
        objItem As MenuItem, _
        obJParent As Object
    
    On Error Resume Next
    
    Set objAccels = Menu.Accelerators
    If objAccels Is Nothing Then Exit Sub
    
    objAccels.Clear
    
    For Each objItem In Menu
        
        If (objItem.Visible = True) Then
            Set objItem.Accelerator.Table = objAccels
            objAccels.Col.Add objItem.Accelerator, "_H" + Hex(objItem.ItemId)
            
            Set obJParent = objItem.Submenu
            UpdateTables obJParent
            
            Set obJParent = Nothing
        Else
            Set objItem.Accelerator.Table = Nothing
        End If
        
    Next objItem
    
    Set objAccels = Nothing
    
End Sub

Public Sub RecreateMenu(Menu As Object)

    Dim hMenu As Long, _
        ItemId As Long, _
        cbFlags As Long
        
    Dim hWnd As Long
    
    Dim tCol As Collection
        
    Dim varItem As Object
    
    If Not TypeOf Menu Is Menubar Then
        
        Set varItem = Menu.Sidebar
        
        If Not varItem Is Nothing Then
            RefreshItem varItem.ItemId
            Set varItem = Nothing
        End If
        
    ElseIf (Menu.Connected = True) Then
        
        hWnd = Menu.hWnd
        hMenu = Menu.hMenu
        
        SetMenu hWnd, 0&
        DrawMenuBar hWnd
        
    End If
    
    For Each varItem In Menu
        RefreshItem varItem.ItemId, False
    Next varItem
            
    ReindexMenu Menu
    
    If (TypeOf Menu Is Menubar) Then
        If (Menu.Connected = True) Then
            SetMenu hWnd, hMenu
        Else
            hWnd = Menu.hWnd
        End If
        
        If (hWnd <> 0&) Then
            DrawMenuBar hWnd
        End If
            
    End If
        
    
End Sub

Public Function CopyMenu_API(ByVal hMenu As Long, ByVal Destination As Submenu, Optional ByVal Subclass As Boolean, _
    Optional ByVal fSaveOrigId As Boolean, Optional ByVal fTranslate As TranslateItemDataConstants) As Boolean
    
    Dim i As Long, _
        j As Long
        
    Dim iState As Long, _
        newPic As Long
        
    Dim cbFlags As Long, _
        lpInfo As MENUITEMINFO
        
    Dim sCaption As String, _
        sAccel As String
        
    Dim a As Long, _
        b As Long
        
    Dim Item As MenuItem, _
        Col As Collection
        
    Dim destMenu As Long
        
    If Destination Is Nothing Then Exit Function
        
    j = GetMenuItemCount(hMenu) - 1&
    
    cbFlags = MF_BYPOSITION
    
    Set Col = Destination.Col
    
    If (Subclass = True) Then
        destMenu = hMenu
        
        Destination.Destroy
        Destination.hMenu = hMenu
    Else
        destMenu = Destination.hMenu
        
        If (destMenu = 0&) Then
            Destination.Create
            destMenu = Destination.hMenu
        End If
    End If
    
    For i = 0& To j
        
        lpInfo.fMask = (MIIM_CHECKMARKS Or _
                        MIIM_TYPE Or _
                        MIIM_STATE Or _
                        MIIM_SUBMENU Or _
                        MIIM_ID)
                    
        sCaption = ""
        GetMenuItemInfo_API hMenu, i, cbFlags, lpInfo, sCaption
        
        If ((lpInfo.fType And CopyableMask) <> 0&) Or (lpInfo.fType = MFT_STRING) Then
            With lpInfo
                
                If Not Item Is Nothing Then
                    Set Item.NextItem = New MenuItem
                    Set Item.NextItem.PrevItem = Item
                    Set Item = Item.NextItem
                Else
                
                    Set Item = New MenuItem
                End If
                
                Set Item.Parent = Destination
                Set Item.Accelerator.Table = Destination.Accelerators
                
                Item.Index = i
                                            
                If (fTranslate <> tdNone) Then
                    Item.ItemData = lpInfo.dwItemData
                End If
                            
                If (fSaveOrigId = True) Then
                    Item.Tag = "" & lpInfo.wID
                End If
                
                If (Subclass = True) Then
                    Item.ItemId = lpInfo.wID
                Else
                    lpInfo.wID = Item.ItemId
                End If
                
                Item.Key = "_H" + Hex(lpInfo.wID)
                
                Col.Add Item, Item.Key
                
                If (.fType And MFT_RIGHTJUSTIFY) Then
                    Item.Visual.TextAlign = taRight
                End If
            
                If (.fState And MFS_DEFAULT) Then
                    Item.Default = True
                End If
                
                If (.fState And MFS_DISABLED) Then
                    Item.Enabled = False
                End If
                
                If (.fType And MFT_RADIOGROUP) Then
                    Item.RadioGroup = True
                End If
                
                If (.fState And MFS_CHECKED) Then
                    Item.Checked = True
                End If
                                                    
                If (sCaption <> "") Then
                    a = InStr(1, sCaption, vbTab)
                    
                    If (a <> 0&) Then
                        sAccel = Mid(sCaption, a + 1)
                        sCaption = Mid(sCaption, 1, a - 1)
                    
                        Item.Accelerator.AccelWord = sAccel
                    End If
                    
                    Item.Caption = sCaption
                End If
                
                If (.fType And MFT_SEPARATOR) Then
                    Item.Separator = True
                    Item.SeparatorType = mstNormal
                    
                    Item.Caption = "-"
                ElseIf (.fType And MFT_MENUBARBREAK) Then
                    Item.Separator = True
                    Item.SeparatorType = mstBarBreak
                ElseIf (.fType And MFT_MENUBREAK) Then
                    Item.Separator = True
                    Item.SeparatorType = mstBreak
                End If
                
                If (lpInfo.hbmpUnchecked <> 0&) Then
                    newPic = CopyImage(lpInfo.hbmpUnchecked, IMAGE_BITMAP, 16, 16, LR_COPYFROMRESOURCE)
                    Set Item.Picture = GetOlePicture(newPic, IMAGE_BITMAP, True)
                End If
                
                If (lpInfo.hbmpChecked <> 0&) Then
                    newPic = CopyImage(lpInfo.hbmpChecked, IMAGE_BITMAP, 16, 16, LR_COPYFROMRESOURCE)
                    Set Item.CheckedPicture = GetOlePicture(newPic, IMAGE_BITMAP, True)
                    Item.Visual.SelectionStyle = mssNoCheckBevel
                End If
                
                
                If (Subclass = False) Then
                    AppendMenu destMenu, MF_OWNERDRAW, Item.ItemId, 0&
                Else
                    DeleteMenu hMenu, i, MF_BYPOSITION
                    InsertMenu hMenu, i, MF_BYPOSITION + MF_OWNERDRAW, Item.ItemId, 0&
                End If
                
                If (lpInfo.hSubMenu <> 0&) Then
                    CopyMenu_API lpInfo.hSubMenu, Item.Submenu, False, fSaveOrigId
                    
                    lpInfo.fMask = MIIM_SUBMENU
                    lpInfo.hSubMenu = Item.Submenu.hMenu
                    
                    SetMenuItemInfo_API hMenu, Item.ItemId, MF_BYCOMMAND, lpInfo
                End If
                
                Item.Unfreeze
                
                If (Subclass = True) Then
                    Item.SetCtrlState
                Else
                    Item.SetCtrlState (True)
                End If
                
            End With
        End If
    
    Next i

    Set Item = Nothing
    Set Col = Nothing

    ReindexMenu Destination
    
End Function


Public Function CopySystemMenu_API(ByVal hWnd As Long, ByVal Destination As SystemMenu, Optional ByVal fRevert As Boolean, Optional ByVal fTranslate As TranslateItemDataConstants) As Boolean
    Dim hMenu As Long, _
        i As Long, _
        j As Long
        
    Dim iState As Long
        
    Dim cbFlags As Long, _
        lpInfo As MENUITEMINFO
        
    Dim sCaption As String, _
        sAccel As String
        
    Dim a As Long, _
        b As Long
        
    Dim Item As MenuItem, _
        Col As Collection
        
    If Destination Is Nothing Then Exit Function
        
    If (fRevert = True) Then
        GetSystemMenu hWnd, fRevert
    End If
    
    hMenu = GetSystemMenu(hWnd, 0&)
    j = GetMenuItemCount(hMenu) - 1&
    
    cbFlags = MF_BYPOSITION
    
    Destination.Clear
    Set Col = Destination.Col
    
    Destination.hMenu = hMenu
    
    
    Destination.hWnd = hWnd
    
    For i = 0& To j
        
        lpInfo.fMask = (MIIM_CHECKMARKS Or _
                        MIIM_TYPE Or _
                        MIIM_STATE Or _
                        MIIM_SUBMENU Or _
                        MIIM_ID)
                    
        GetMenuItemInfo_API hMenu, i, cbFlags, lpInfo, sCaption
        
        With lpInfo
            
            If Not Item Is Nothing Then
                Set Item.NextItem = New MenuItem
                Set Item.NextItem.PrevItem = Item
                Set Item = Item.NextItem
            Else
            
                Set Item = New MenuItem
            End If
            
            Set Item.Parent = Destination
            Set Item.Accelerator.Table = Destination.Accelerators
                            
            If (fTranslate <> tdNone) Then
                Item.ItemData = lpInfo.dwItemData
            End If
                            
            Item.Index = i
            
            RemoveMenu hMenu, i, MF_BYPOSITION
            InsertMenu hMenu, i, MF_BYPOSITION, lpInfo.wID, 0&
            
            Item.ItemId = lpInfo.wID
            
            Item.Key = "_H" + Hex(lpInfo.wID)
            Item.ParentType = mtcSysmenu
            
            Col.Add Item, Item.Key
            
            If (.fType And MFT_RIGHTJUSTIFY) Then
                Item.Visual.TextAlign = taRight
            End If
        
            If (.fType And MFT_SEPARATOR) Then
                Item.Separator = True
                Item.SeparatorType = mstNormal
            ElseIf (.fType And MFT_MENUBARBREAK) Then
                Item.Separator = True
                Item.SeparatorType = mstBarBreak
            ElseIf (.fType And MFT_MENUBREAK) Then
                Item.Separator = True
                Item.SeparatorType = mstBreak
            End If
            
            If (.fState And MFS_DEFAULT) Then
                Item.Default = True
            End If
            
            If (.fState And MFS_DISABLED) Then
                Item.Enabled = False
            End If
            
            If (.fType And MFT_RADIOGROUP) Then
                Item.RadioGroup = True
            End If
            
            If (.fState And MFS_CHECKED) Then
                Item.Checked = True
            End If
                                                                                                        
            If (lpInfo.hSubMenu <> 0&) Then
                CopyMenu_API lpInfo.hSubMenu, Item.Submenu, False, True
                
                lpInfo.fMask = MIIM_SUBMENU
                lpInfo.hSubMenu = Item.Submenu.hMenu
                
                SetMenuItemInfo_API hMenu, lpInfo.wID, MF_BYCOMMAND, lpInfo
            End If
            
            iState = GetWndState(hWnd)
            
            '' Set our own default system menu bitmaps
            '' set the system-menu state to match reality
            
            Select Case i
                Case 0:  ''Restore
                    Item.Caption = "&Restore"
                    Set Item.Picture = LoadResPicture(ResImage_Restore, vbResBitmap)
                                        
                    If (iState = 0&) Then Item.Enabled = False
                    
                Case 1:  ''Move
                    Item.Caption = "&Move"
                
                Case 2:  ''Size
                    Item.Caption = "&Size"
                                
                Case 3:  ''Minimize
                    
                    Item.Caption = "Mi&nimize"
                    Set Item.Picture = LoadResPicture(ResImage_Minimize, vbResBitmap)
                                        
                    If (iState = 2&) Then Item.Enabled = False
                
                Case 4:  ''Maximize
                    
                    Item.Caption = "Ma&ximize"
                    Set Item.Picture = LoadResPicture(ResImage_Maximize, vbResBitmap)
                    
                    If (iState = 1&) Then Item.Enabled = False
                    
                Case 5:  ''Separator
                    Item.Caption = "-"
                
                Case 6:  ''Close
                    Item.Caption = "&Close"
                    Set Item.Picture = LoadResPicture(ResImage_Close, vbResBitmap)
                    Item.Accelerator.AccelWord = "Alt+F4"
                    Item.Default = True
            
            
            End Select
            
            Item.Unfreeze
            Item.SetCtrlState
        
        End With
    
    Next i
    
    Set Item = Nothing
    Set Col = Nothing
    
    ReindexMenu Destination
        
End Function

Public Function CopyMenubar_API(ByVal hMenu As Long, ByVal Destination As Menubar, Optional ByVal Subclass As Boolean, Optional ByVal fSaveOrigId As Boolean, Optional ByVal fTranslate As TranslateItemDataConstants)
    Dim i As Long, _
        j As Long
        
    Dim iState As Long, _
        deadInfo As MENUITEMINFO
        
    Dim cbFlags As Long, _
        lpInfo As MENUITEMINFO
        
    Dim sCaption As String, _
        sAccel As String
        
    Dim a As Long, _
        b As Long
        
    Dim Item As MenuItem, _
        Col As Collection
        
    Dim destMenu As Long
        
    If Destination Is Nothing Then Exit Function
        
    j = GetMenuItemCount(hMenu) - 1&
    
    cbFlags = MF_BYPOSITION
    
    Set Col = Destination.Col
    
    If (Subclass = True) Then
        destMenu = hMenu
        
        Destination.Destroy
        Destination.hMenu = hMenu
    Else
        destMenu = Destination.hMenu
        If (destMenu = 0&) Then
            Destination.Create
            destMenu = Destination.hMenu
        End If
    End If
    
    For i = 0& To j
        
        CopyMemory lpInfo, deadInfo, Len(lpInfo)
        sCaption = ""
        
        lpInfo.fMask = (MIIM_CHECKMARKS Or _
                        MIIM_TYPE Or _
                        MIIM_STATE Or _
                        MIIM_SUBMENU Or _
                        MIIM_ID)
                    
        GetMenuItemInfo_API hMenu, i, cbFlags, lpInfo, sCaption
        
        If ((lpInfo.fType And CopyableMask) <> 0&) Or (lpInfo.fType = MFT_STRING) Then
            With lpInfo
                
                If Not Item Is Nothing Then
                    Set Item.NextItem = New MenuItem
                    Set Item.NextItem.PrevItem = Item
                    Set Item = Item.NextItem
                Else
                
                    Set Item = New MenuItem
                End If
                
                Set Item.Parent = Destination
                Set Item.Accelerator.Table = Destination.Accelerators
                                
                Item.Visual.TextAlign = taCenter
                
                Item.ParentType = mtcMenubar
                
                Item.Index = i
                            
                If (fTranslate <> tdNone) Then
                    Item.ItemData = lpInfo.dwItemData
                End If
                                                
                If (fSaveOrigId = True) Then
                    Item.Tag = "" & lpInfo.wID
                End If
                
                If (Subclass = True) Then
                    Item.ItemId = lpInfo.wID
                End If
                
                Item.Key = "_H" + Hex(Item.ItemId)
                
                Col.Add Item, Item.Key
                
                If (.fType And MFT_RIGHTJUSTIFY) Then
                    Item.Visual.TextAlign = taRight
                End If
            
                If (.fType And MFT_SEPARATOR) Then
                    Item.Separator = True
                    Item.SeparatorType = mstNormal
                ElseIf (.fType And MFT_MENUBARBREAK) Then
                    Item.Separator = True
                    Item.SeparatorType = mstBarBreak
                ElseIf (.fType And MFT_MENUBREAK) Then
                    Item.Separator = True
                    Item.SeparatorType = mstBreak
                End If
                
                If (.fState And MFS_DEFAULT) Then
                    Item.Default = True
                End If
                
                If (.fState And MFS_DISABLED) Then
                    Item.Enabled = False
                End If
                
                If (.fType And MFT_RADIOGROUP) Then
                    Item.RadioGroup = True
                End If
                
                If (.fState And MFS_CHECKED) Then
                    Item.Checked = True
                End If
                                                    
                If (sCaption <> "") Then
                    a = InStr(1, sCaption, vbTab)
                    
                    If (a <> 0&) Then
                        sAccel = Mid(sCaption, a + 1)
                        sCaption = Mid(sCaption, 1, a - 1)
                    
                        Item.Accelerator.AccelWord = sAccel
                    End If
                    
                    Item.Caption = sCaption
                End If
                
                If (Subclass = False) Then
                    AppendMenu destMenu, MF_OWNERDRAW, Item.ItemId, 0&
                Else
                    DeleteMenu hMenu, i, MF_BYPOSITION
                    InsertMenu hMenu, i, MF_BYPOSITION + MF_OWNERDRAW, Item.ItemId, 0&
                End If
                
                If (lpInfo.hSubMenu <> 0&) Then
                    CopyMenu_API lpInfo.hSubMenu, Item.Submenu, False, fSaveOrigId
                End If
                
                Item.Unfreeze
                
                If (Subclass = True) Then
                    Item.SetCtrlState
                Else
                    Item.SetCtrlState (True)
                End If
                
            End With
        End If
    
    Next i


    Set Item = Nothing
    Set Col = Nothing
    
    ReindexMenu Destination
    
End Function

Public Function IsParent(MenuComp As Object, MenuRef As Object) As Boolean
    On Error Resume Next
    
    Dim obJParent As Object
    
    Set obJParent = MenuRef.Parent
    
    IsParent = False
    
    Do While Not obJParent Is Nothing
        If obJParent Is MenuComp Then
            IsParent = True
            Exit Do
        End If
        
        Set obJParent = obJParent.Parent
    Loop
    
    Set obJParent = Nothing
        
End Function

'' Is this a system command ID?

Public Function IsSysCmd(ByVal nCmd As Long) As Long
    Dim wCmd As Long
    
    wCmd = (nCmd And &HFFF0&)
    
    Select Case wCmd
    
        Case SC_ARRANGE, SC_CLOSE, SC_HOTKEY, SC_HSCROLL, SC_ICON, _
             SC_MAXIMIZE, SC_MINIMIZE, SC_MOUSEMENU, SC_MOVE, _
             SC_NEXTWINDOW, SC_PREVWINDOW, SC_RESTORE, SC_SCREENSAVE, SC_SIZE, _
             SC_TASKLIST, SC_VSCROLL, SC_ZOOM
    
            IsSysCmd = wCmd
            
        Case Else
            
            IsSysCmd = 0&
            
    End Select
    
End Function

Public Function SendCommand(Menu As Object, ByVal wParam As Long, ByVal uMsg As Long)

    Dim objSubmenu As Submenu, _
        objMenubar As Menubar, _
        objSysMenu As SystemMenu

    If TypeOf Menu Is Menubar Then
        Set objMenubar = Menu
        objMenubar.ExecCmd wParam, uMsg
        Set objMenubar = Nothing
    ElseIf TypeOf Menu Is Submenu Then
        Set objSubmenu = Menu
        objSubmenu.ExecCmd wParam, uMsg
        Set objSubmenu = Nothing
        
    ElseIf TypeOf Menu Is SystemMenu Then
        Set objSysMenu = Menu
        objSysMenu.ExecCmd wParam, uMsg
        Set objSysMenu = Nothing
    Else
        '' Pray for the best
        On Error Resume Next
        Menu.ExecCmd wParam, uMsg
    End If
    
End Function

Public Function SetImageMax(Menu As Object, ByVal lMax As Long, ByVal rMax As Long)

    Dim objSubmenu As Submenu, _
        objMenubar As Menubar, _
        objSysMenu As SystemMenu

    If TypeOf Menu Is Menubar Then
        Set objMenubar = Menu
        objMenubar.SetImageMax lMax, rMax
        Set objMenubar = Nothing
    ElseIf TypeOf Menu Is Submenu Then
        Set objSubmenu = Menu
        objSubmenu.SetImageMax lMax, rMax
        Set objSubmenu = Nothing
        
    ElseIf TypeOf Menu Is SystemMenu Then
        Set objSysMenu = Menu
        objSysMenu.SetImageMax lMax, rMax
        Set objSysMenu = Nothing
    Else
        '' Pray for the best
        On Error Resume Next
        Menu.SetImageMax lMax, rMax
    End If
    
End Function

Public Function MenuMaxImageWidth(Menu As Object, Optional RightWidth) As Long
    Dim objItem As MenuItem, _
        fPicSize As SIZEAPI, _
        fPicComp As SIZEAPI
    
    Dim cxLeft As Long, _
        cxRight As Long

    On Error Resume Next
    
    For Each objItem In Menu
        
        If (Not objItem.Picture Is Nothing) Or _
            ((objItem.Checked = True) And (Not objItem.CheckedPicture Is Nothing)) Then
            
            If (objItem.Visual.ScaleImages = True) Then
                fPicSize.cx = (IconicPad + objItem.Visual.ImageScaleWidth)
                fPicSize.cy = (IconicPad + objItem.Visual.ImageScaleHeight)
            Else
                If Not objItem.Picture Is Nothing Then
                    fPicSize.cx = ScaleTool.DeviceTranslateX(objItem.Picture.Width, nHiMetric, nPixels)
                    fPicSize.cy = ScaleTool.DeviceTranslateY(objItem.Picture.Height, nHiMetric, nPixels)
                End If
                
                If Not objItem.CheckedPicture Is Nothing Then
                    fPicComp.cx = ScaleTool.DeviceTranslateX(objItem.CheckedPicture.Width, nHiMetric, nPixels)
                    fPicComp.cy = ScaleTool.DeviceTranslateY(objItem.CheckedPicture.Height, nHiMetric, nPixels)
                End If
                
                If (fPicComp.cx > fPicSize.cx) Then fPicSize.cx = fPicComp.cx
                fPicSize.cx = IconicPad + fPicSize.cx
                
                If (IconicSquare > fPicSize.cx) Then
                    fPicSize.cx = IconicSquare
                End If
                
                If (fPicComp.cy > fPicSize.cy) Then fPicSize.cy = fPicComp.cy
                fPicSize.cy = IconicPad + fPicSize.cy
                
                If (IconicSquare > fPicSize.cy) Then
                    fPicSize.cy = IconicSquare
                End If
                    
            End If
            
            If (objItem.Visual.TextAlign = taRight) Or (objItem.RightAlign = True) Or _
                (objItem.RightToLeft = True) Then
                
                If (cxRight < fPicSize.cx) Then cxRight = fPicSize.cx
            Else
                If (cxLeft < fPicSize.cx) Then cxLeft = fPicSize.cx
            End If
            
        End If
    
    Next objItem

    If IsMissing(RightWidth) = False Then
        RightWidth = cxRight
    End If
    
    MenuMaxImageWidth = cxLeft
    SetImageMax Menu, cxLeft, cxRight

End Function

''' Get a new control ID for the menu interface from the global pool

Public Function GetNewCtrlId() As Long
    
    If g_CtrlIdNext = 0& Then
        g_CtrlIdNext = &H1248&
    End If
                
    GetNewCtrlId = g_CtrlIdNext
    
    '' Increment the control ID.
    g_CtrlIdNext = g_CtrlIdNext + 1&
            
End Function

'' Get the system default menu font

Public Function GetSysMenuFont() As StdFont
    Dim lpNCMetrics As NONCLIENTMETRICS
    
    lpNCMetrics.cbSize = Len(lpNCMetrics)
    SystemParametersInfo SPI_GETNONCLIENTMETRICS, Len(lpNCMetrics), lpNCMetrics, 0&
    
    Set GetSysMenuFont = GetOleFont(lpNCMetrics.lfMenuFont)
        
End Function


'' Get Windows version information and set the length of
'' Menu item's info.

Public Function GetWinVersion() As WindowsVersionConstants

    Dim pData As OSVERSIONINFO, m_ItemInfo As MENUITEMINFO
    
    pData.OSVersionInfoSize = Len(pData)
    
    GetVersion pData
    
    g_WinVersion = (pData.MajorVersion * 100) + pData.MinorVersion

    Select Case g_WinVersion
    
        Case Windows2000:
            g_InfoSize = Len(m_ItemInfo)
        
        Case Else
            g_InfoSize = Len(m_ItemInfo) - 4
    
    End Select

End Function


'' Get a printed string containing information about the current operating system.
Public Function WinVerString() As String

    Dim pData As OSVERSIONINFO
    Dim v_Str As String
    
    pData.OSVersionInfoSize = Len(pData)
    
    GetVersion pData
        
    g_WinVersion = (pData.MajorVersion * 100) + pData.MinorVersion
    
    v_Str = "Windows "
    
    Select Case g_WinVersion
    
        Case Windows95:
            v_Str = v_Str + "95"
        Case Windows98:
            v_Str = v_Str + "95"
        Case WindowsNT:
            v_Str = v_Str + "NT"
        Case WindowsME:
            v_Str = v_Str + "Millennium"
        Case Windows2000:
            v_Str = v_Str + "2000"
        Case Else
            If (g_WinVersion < 600) And (g_WinVersion > 500) Then
                v_Str = v_Str + "XP"
            Else
                If pData.PlatformId = VER_PLATFORM_WIN32_NT Then
                    v_Str = v_Str + "NT (Unknown)"
                Else
                    v_Str = v_Str + "32-bit (Unknown)"
                End If
            End If
    End Select
    
    With pData
    
        If InStr(1, .CSDVersion, Chr(0&)) Then
            .CSDVersion = Mid(.CSDVersion, 1, InStr(1, .CSDVersion, Chr(0&)) - 1)
        End If
        v_Str = v_Str + " " + .CSDVersion
    
    End With
    
    WinVerString = v_Str
    
End Function

Public Function GetWndState(ByVal hWnd As Long)
    Dim lpInfo As WINDOWPLACEMENT
    
    lpInfo.Length = Len(lpInfo)
    GetWindowPlacement hWnd, lpInfo
    
    Select Case lpInfo.showCmd
    
        Case SW_MINIMIZE
            GetWndState = 2&
        
        Case SW_MAXIMIZE
            GetWndState = 1&
    
        Case Else
            GetWndState = 0&
    
    End Select
    
End Function

Public Sub ChangeActiveMenuSet(ByVal MenuSet As Menus, Optional ByVal MDIConnection As Object)
    Dim varObj As Object
    
    Dim aWnd() As MENUKEY, _
        x As Long, _
        i As Long
    
    If MenuSet Is Nothing Then Exit Sub
    
    Dim connectKey As String
    
    On Error Resume Next
    
    If Not MDIConnection Is Nothing Then
        connectKey = MDIConnection.Menubar.Key
    End If
    
    For Each varObj In g_MenuCol
                
        If (TypeOf varObj Is Menubar) Then
        
            ReDim Preserve aWnd(0 To x)
            aWnd(x).hWnd = varObj.hWnd
            aWnd(x).Key = varObj.Key
        
            x = x + 1
            
        End If
        
        varObj.Destroy
    Next varObj
            
    g_MenuCol.Active = False
    
    Set g_MenuCol = Nothing
    Set g_MenuCol = MenuSet
    
    g_MenuCol.Active = True
    
    x = x - 1&
    For i = 0& To x
    
        Set varObj = g_MenuCol(aWnd(i).Key)
        
        If (Not varObj Is Nothing) And (TypeOf varObj Is Menubar) Then
            varObj.SetWindowHandle aWnd(i).hWnd
        End If
        
        Set varObj = Nothing
    
    Next i

    If (connectKey <> "") Then
        Set varObj = g_MenuCol(connectKey)
        
        If (Not varObj Is Nothing) And (TypeOf varObj Is Menubar) Then
            MDIConnection.Connect varObj
        End If
    End If
    
    Erase aWnd
    
    
End Sub


Public Sub AddToMenuSet(ByVal Menu As Object)
    Dim sKey As String
    
    If (Menu.Key <> "") Then
        sKey = Menu.Key
    Else
        sKey = "_H" + Hex(Menu.hMenu)
    End If
    
    Menu.Key = sKey
    
    If g_MenuCol Is Nothing Then Set g_MenuCol = New Menus
    
    If g_MenuCol(sKey) Is Nothing Then
            
        g_MenuCol.Add Menu, sKey
    
    ElseIf Not g_MenuCol(sKey) Is Menu Then
    
        g_MenuCol.Remove sKey
        g_MenuCol.Add Menu, sKey
    
    End If

End Sub

Public Sub RemoveFromMenuSet(ByVal Menu As Object)
    Dim sKey As String
    
    If (Menu.Key <> "") Then
        sKey = Menu.Key
    Else
        sKey = "_H" + Hex(Menu.hMenu)
    End If

    If Not g_MenuCol(sKey) Is Nothing Then
        g_MenuCol.Remove sKey
    End If
    
    If g_MenuCol.Count = 0& Then Set g_MenuCol = Nothing
    
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





