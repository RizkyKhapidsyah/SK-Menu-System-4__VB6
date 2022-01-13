VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SubMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"MenuItem"
Attribute VB_Ext_KEY = "Member0" ,"MenuItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
''' GUInerd Standard Menu System
''' Version 4.1

''' Objects/API Dll

''' Submenu Class Object


''' ****************************************************************************
''' *** Check the end of the file for copyright and distribution information ***
''' ****************************************************************************

Option Explicit

Private WithEvents m_Sidebar As Sidebar
Attribute m_Sidebar.VB_VarHelpID = -1

Private m_ItemId As Long

Private m_WindowList As Boolean

Private m_Key As String

Private m_hMenu As Long

Private m_hWnd As Long

Private m_Col As Collection

Private m_Accelerators As Accelerators

Private m_Font As StdFont

Private m_Parent As Object

Private m_RightToLeft As Boolean

Private m_LargeIcons As Boolean

Private m_AutoDismiss As Boolean

Private m_MaxHeight As Long

'' Max image widths that are recorded on refresh
Private m_miwLeft As Long

Private m_miwRight As Long

Public Event Initialize()

Public Event UserCommand(ByVal Item As MenuItem)

Public Event ItemRightClick(ByVal Item As MenuItem)

Public Event SubmenuInitialize(ByVal Submenu As Submenu)

Public Event SidebarClick()

Public Event SidebarRightClick()

'''' Public (and Public Read-Only) Properties

Public Property Get hMenu() As Long
    hMenu = m_hMenu
End Property

Friend Property Let hMenu(ByVal vData As Long)
    
    On Error Resume Next
    
    If m_hMenu <> 0& Then Destroy
    
    m_hMenu = vData
    
    If (m_hMenu <> 0&) Then
        AddToMenuSet Me
    End If
    
End Property

Public Property Get hWnd() As Long
    hWnd = m_hWnd
End Property

Public Property Get ItemId() As Long
    ItemId = m_ItemId
End Property

Friend Property Let ItemId(ByVal vData As Long)
    m_ItemId = vData
End Property

Public Property Get Key() As String
    Key = m_Key
End Property

Public Property Let Key(ByVal vData As String)
    m_Key = vData
End Property

Public Property Get Accelerators() As Accelerators
    If Not m_Parent Is Nothing Then
        Set Accelerators = m_Accelerators
    End If
End Property

Public Property Get Font() As StdFont
    Set Font = m_Font
End Property

Public Property Set Font(IFont As StdFont)

    Set m_Font = IFont
    On Error Resume Next
    
    If m_Parent Is Nothing Then Exit Property
    If TypeOf m_Parent Is Menubar Then
        DrawMenuBar m_Parent.hWnd
    End If
    
End Property

Public Property Get Sidebar() As Sidebar
    Set Sidebar = m_Sidebar
End Property

Public Property Get WindowList() As Boolean
    WindowList = m_WindowList
End Property

Friend Property Let WindowList(ByVal vData As Boolean)
    m_WindowList = vData
End Property

Public Property Get RightToLeft() As Boolean
    RightToLeft = m_RightToLeft
End Property

Public Property Let RightToLeft(ByVal vData As Boolean)
    m_RightToLeft = vData
End Property

'' Volatile properties (read-only)
Public Property Get ImageMax_Left() As Long
    ImageMax_Left = m_miwLeft
End Property

Public Property Get ImageMax_Right() As Long
    ImageMax_Right = m_miwRight
End Property

'' Friend Sub sets data
Friend Sub SetImageMax(ByVal imLeft As Long, ByVal imRight As Long)

    m_miwLeft = imLeft
    m_miwRight = imRight
    
End Sub

Public Property Get MaxHeight() As Long
    Dim lpInfo As MENUINFO

    If (m_hMenu <> 0&) Then
        lpInfo.cbSize = Len(lpInfo)
        lpInfo.fMask = MIM_MAXHEIGHT
        
        GetMenuInfo m_hMenu, lpInfo
        
        m_MaxHeight = lpInfo.cyMax
    End If

    MaxHeight = m_MaxHeight

End Property

Public Property Let MaxHeight(ByVal vData As Long)
    Dim lpInfo As MENUINFO
    
    m_MaxHeight = vData

    If (m_hMenu <> 0&) Then
        lpInfo.cbSize = Len(lpInfo)
        
        lpInfo.cyMax = m_MaxHeight
        lpInfo.fMask = MIM_MAXHEIGHT
        
        SetMenuInfo m_hMenu, lpInfo
    End If
    
End Property

Public Property Get AutoDismiss() As Boolean
    Dim lpInfo As MENUINFO

    If (m_hMenu <> 0&) Then
        lpInfo.cbSize = Len(lpInfo)
        lpInfo.fMask = MIM_STYLE
        
        GetMenuInfo m_hMenu, lpInfo
        
        m_AutoDismiss = CBool((lpInfo.dwStyle And MNS_AUTODISMISS) <> 0&)
                
    End If

    AutoDismiss = m_AutoDismiss

End Property

Public Property Let AutoDismiss(ByVal vData As Boolean)
    Dim lpInfo As MENUINFO
    
    m_AutoDismiss = vData

    If (m_hMenu <> 0&) Then
        lpInfo.cbSize = Len(lpInfo)
        
        lpInfo.fMask = MIM_STYLE
        lpInfo.dwStyle = (m_AutoDismiss And MNS_AUTODISMISS)
        
        SetMenuInfo m_hMenu, lpInfo
    End If
    
End Property

Public Property Get LargeIcons() As Boolean
    LargeIcons = m_LargeIcons
End Property

Public Property Let LargeIcons(ByVal vData As Boolean)
    If (m_LargeIcons <> vData) Then
        m_LargeIcons = vData
        RecreateMenu Me
    End If
End Property

Public Property Get AnyItemVisible() As Boolean
    Dim objItem As MenuItem
    
    If m_Col.Count > 0& Then
        For Each objItem In m_Col
            If objItem.Visible = True Then
                AnyItemVisible = True
                Exit Property
            End If
        Next objItem
    End If
            
    AnyItemVisible = False
    
End Property

Public Property Get Parent() As Object
    Set Parent = m_Parent
End Property

Friend Property Set Parent(ByVal vData As Object)
    On Error Resume Next
    
    If Not vData Is Nothing Then
        If (TypeOf vData Is Menubar) Or (TypeOf vData Is Submenu) Then
            Set m_Parent = vData
            Set m_Font = m_Parent.Font
        End If
    End If
End Property


Public Property Get TopMostParent() As Object
    On Error Resume Next

    If m_Parent Is Nothing Then
        Set TopMostParent = Me
    Else
        Set TopMostParent = m_Parent.TopMostParent
    End If
End Property


'''' Collection and item management properties and methods

Public Function Add(ByVal Caption As String, Optional Picture As StdPicture, Optional ByVal sKey As String, Optional InsertAfter, Optional ByVal hWndList As Long = 0&) As MenuItem
    On Error Resume Next
    
    Dim lpInfo As MENUITEMINFO, _
        cbFlags As Long, _
        varStr As String
    
    Dim objNewItem As New MenuItem, _
        objBefore As MenuItem
    
    If m_Col Is Nothing Then Set m_Col = New Collection
    
    Dim hTop As Long
        
    If (IsMissing(InsertAfter) = False) And (hWndList = 0&) Then
        If VarType(InsertAfter) = vbString Then
            Set objBefore = Item(InsertAfter)
        ElseIf IsObject(InsertAfter) Then
            Set objBefore = InsertAfter
        Else
            Set objBefore = g_MenuCol.SearchRecursive(Me, InsertAfter)
        End If
        
        If Not objBefore.Parent Is Me Then
            If TypeOf objBefore.Parent Is Submenu Then
                Set Add = objBefore.Parent.Add(Caption, Picture, sKey, objBefore, hWndList)
            Else
                Set Add = objBefore.Parent.Add(Caption, Picture, sKey, objBefore)
            End If
            
            Exit Function
        End If
    End If
    
    Set objNewItem = New MenuItem
            
    If m_hMenu = 0& Then
        Create
    End If
                
        ''' If this is a sub menu, we need to tell the parent
        ''' menu that this item now links to a sub menu.
    
    If (Not m_Parent Is Nothing) Then
        lpInfo.fMask = MIIM_SUBMENU
            
        lpInfo.wID = m_ItemId
        lpInfo.hSubMenu = m_hMenu
            
        SetMenuItemInfo_API m_Parent.hMenu, m_ItemId, False, lpInfo
    End If
        
    Set objNewItem.Parent = Me
    Set objNewItem.Accelerator.Table = m_Accelerators
    
    cbFlags = (MF_OWNERDRAW + MF_BYCOMMAND)
    
    objNewItem.RightToLeft = m_RightToLeft
    
    
    '' If the handle to a window list has been specified,
    '' then this item gets added to the window-list collection.
    
    If (hWndList <> 0&) And (m_WindowList = True) Then
                
        If g_WindowList Is Nothing Then
            Set g_WindowList = New Collection
        End If
        
        objNewItem.ItemId = hWndList
        objNewItem.Key = "_H" + Hex(hWndList)
        
        objNewItem.Index = m_Col.Count + g_WindowList.Count
        
        g_WindowList.Add objNewItem, objNewItem.Key
        AppendMenu m_hMenu, cbFlags, objNewItem.ItemId, 0&
    
    Else
    
        If sKey = "" Then
            objNewItem.Key = "_H" + Hex(objNewItem.ItemId)
        Else
            objNewItem.Key = sKey
        End If
            
        If Not objBefore Is Nothing Then
                            
            m_Col.Add objNewItem, objNewItem.Key, , objBefore.Index + 1
            
            ReindexMenu Me
            
            InsertMenu m_hMenu, objBefore.ItemId, cbFlags, objNewItem.ItemId, 0&
        
        Else
            If (m_Col.Count <> 0&) Then
                Set objNewItem.PrevItem = m_Col(m_Col.Count)
                Set objNewItem.PrevItem.NextItem = objNewItem
            End If
            
            objNewItem.Index = m_Col.Count
            
            AppendMenu m_hMenu, cbFlags, objNewItem.ItemId, 0&
            m_Col.Add objNewItem, objNewItem.Key
            
        End If
        
    End If
            
    objNewItem.Caption = Caption
    
    If Not Picture Is Nothing Then
        Set objNewItem.Picture = Picture
    End If
    
    '' If the global drawing Style is for OfficeXP, pre-configure
    '' menu items (this is to allow flexibility to remain within the item
    '' as much as possible.  These configurations are suggestions, they can
    '' change, with dramatic results of done tastefully and carefully!)
    
    If (g_MenuCol.MenuDrawStyle = mdsOfficeXP) Then
        objNewItem.Visual.SelectionStyle = (objNewItem.Visual.SelectionStyle Or mssFlat)
    End If
        
    objNewItem.Unfreeze
    objNewItem.SetCtrlState True
    
    RefreshItem m_ItemId
    
    Set Add = objNewItem
    Set objNewItem = Nothing
    
End Function

Public Property Get Item(varIndexKey) As MenuItem
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
        
    On Error Resume Next
    
    
    If m_Col Is Nothing Then Set m_Col = New Collection
    
    Select Case VarType(varIndexKey)
    
        Case vbString
        
            If Not m_Col(varIndexKey) Is Nothing Then
                
                Set Item = m_Col(varIndexKey)
                Exit Property
            
            ElseIf m_WindowList = True Then
                
                If Not g_WindowList(varIndexKey) Is Nothing Then
                    Set Item = g_WindowList(varIndexKey)
                    Exit Property
                End If
            
            End If
        
        Case vbLong, vbInteger, vbByte
        
            If Not m_Col(varIndexKey) Is Nothing Then
                
                Set Item = m_Col(varIndexKey)
                Exit Property
                
            ElseIf m_WindowList = True Then
            
                If varIndexKey <= 0& Then Exit Property
                
                If Not g_WindowList(varIndexKey - m_Col.Count) Is Nothing Then
                    
                    Set Item = g_WindowList(varIndexKey - m_Col.Count)
                    Exit Property
                    
                End If
                
            End If
            
    End Select
                            
    Set Item = Nothing
                
End Property


Public Property Get Submenu(varIndexKey) As Submenu
        
    On Error Resume Next
    Set Submenu = Item(varIndexKey).Submenu
                
End Property

Public Property Get Count() As Long
    On Error Resume Next
    
    
    If m_Col Is Nothing Then Set m_Col = New Collection
    
    If m_WindowList = True Then
        Count = m_Col.Count + g_WindowList.Count
    Else
        Count = m_Col.Count
    End If
    
End Property

Public Sub Remove(varIndexKey As Variant)
    
    
    If m_Col Is Nothing Then Set m_Col = New Collection
    
    Select Case VarType(varIndexKey)
    
        Case vbString
        
            If Not m_Col(varIndexKey) Is Nothing Then
                m_Col.Remove varIndexKey
            ElseIf m_WindowList = True Then
                If Not g_WindowList(varIndexKey) Is Nothing Then
                    g_WindowList.Remove varIndexKey
                End If
            End If
        
        Case vbLong, vbInteger, vbByte
        
            If Not m_Col(varIndexKey) Is Nothing Then
               m_Col.Remove varIndexKey
            ElseIf m_WindowList = True Then
                If varIndexKey <= 0& Then Exit Sub
                If Not g_WindowList(varIndexKey - m_Col.Count) Is Nothing Then
                    g_WindowList.Remove (varIndexKey - m_Col.Count)
                End If
            End If
            
    End Select
    
    ReindexMenu Me
    
End Sub

Public Sub Clear()
    
    Dim objItem As MenuItem, _
        i As Long, _
        j As Long
    
    
    If m_Col Is Nothing Then
        Set m_Col = New Collection
        Exit Sub
    End If
    
    For Each objItem In m_Col
        objItem.Submenu.Destroy
        objItem.Destroy
    Next objItem
    
    Set m_Col = Nothing
    Set m_Col = New Collection
    
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_Col.[_NewEnum]
End Function

Public Property Get Child() As MenuItem
    If (m_Col.Count <> 0&) Then
        Set Child = m_Col(1)
    Else
        Set Child = Nothing
    End If
End Property

'''' End Collection properties


'' Search for an item anywhere within the items and submenus of this menu.

Public Function Search(varIndexKey) As Object
    Set Search = g_MenuCol.SearchRecursive(Me, varIndexKey)
End Function

'' Forces the window-list collection to be reset (if this submenu holds an MDI window list).

Public Sub ClearWindowList()
    Dim objItem As MenuItem
    
    If g_WindowList.Count > 0& Then
        For Each objItem In g_WindowList
            objItem.Destroy
        Next objItem
           
        Set g_WindowList = New Collection
    End If
    
End Sub

'' Take control of any popup menu with a system handle
'' take over drawing and handling of events without
'' destroying the previous menu

Public Sub Subclass(ByVal hMenu As Long, Optional ByVal fTagOriginalIDs As Boolean, Optional ByVal fTranslate As TranslateItemDataConstants)
    On Error Resume Next
    
    If (m_hMenu <> 0&) Then Exit Sub
    
    If Not m_Parent Is Nothing Then
        Set m_Parent = Nothing
    End If
    
    m_hMenu = hMenu
        
    AddToMenuSet Me
    
    CopyMenu_API m_hMenu, Me, True, fTagOriginalIDs, fTranslate
    
End Sub

'' Public interface to internal CopyMenu function
'' copy any menu with a system handle into this
'' menu, as well as possible.

Public Sub CopyMenu(ByVal hMenu As Long, Optional ByVal fTagOriginalIDs As Boolean, Optional ByVal fTranslate As TranslateItemDataConstants)

    CopyMenu_API hMenu, Me, False, fTagOriginalIDs, fTranslate
    
End Sub

'' Show the dropdown/popup menu as a stand-alone popup menu.

Public Sub Popup(Optional ByVal hWnd As Long, Optional ByVal fRecurse As Boolean)
    Dim lpPoint As POINTAPI, _
        i As Long

    Dim lpRect As RECT
        ' dummy rect for ignored variable.
    

    If m_hMenu = 0& Then Exit Sub
    
    GetCursorPos lpPoint
    
    ' no way to track a popup menu without a window to handle
    ' the callbacks.
    
    If (TypeOf TopMostParent Is Menubar) Then
        If (hWnd = 0&) And (TopMostParent.hWnd = 0&) Then
            Exit Sub
        End If
    
    ElseIf hWnd = 0& Then
        Exit Sub
        
    End If
        
    ' if the user chooses to pass an hWnd at the procedure, we process that one
    ' even if the superparent has a window.  To use the superparent window handle,
    ' the developer should call Popup without any arguments.
    
    If Not hWnd = 0& Then
        
        If GetOldWndProc(hWnd) = 0& Then
            AddHandle hWnd, AddressOf MenuWndProc, m_hMenu
            
            i = TrackPopupMenu(m_hMenu, TPM_RETURNCMD + (TPM_RECURSE And fRecurse), lpPoint.x, lpPoint.y, 0&, hWnd, lpRect)
            RemoveHandle hWnd
            
            ExecCmd i, WM_COMMAND
            
        Else
            TrackPopupMenu m_hMenu, (TPM_RECURSE And fRecurse), lpPoint.x, lpPoint.y, 0&, hWnd, lpRect
        End If
        
    Else
    
        ' TrackPopupMenu is the windows API for opening a menu handle
        ' created as a Submenu/popupmenu... as a popup menu.
        
        ' lpRect is ignored.
        
        TrackPopupMenu m_hMenu, 0&, lpPoint.x, lpPoint.y, 0&, TopMostParent.hWnd, lpRect
    
    End If
    
End Sub

'''' Create a new popup menu and add a reference to this object
'''' in the global menu set

Public Sub Create()
    On Error Resume Next
   
    If (m_hMenu <> 0&) Then Exit Sub
    
    m_hMenu = CreatePopupMenu
    
    If m_Col Is Nothing Then Set m_Col = New Collection
    
    AddToMenuSet Me
    
End Sub

'''' remove reference in global menu set
'''' destroy all items (and submenus), and destroy the popup menu

Public Sub Destroy()
    On Error Resume Next

    Dim objItem As MenuItem
    
    If Not m_Sidebar Is Nothing Then
        m_Sidebar.Destroy
        Set m_Sidebar = Nothing
    End If
    
    For Each objItem In m_Col
        objItem.Submenu.Destroy
        objItem.Destroy
    Next objItem
    
    If (m_hMenu <> 0&) Then
        
        RemoveFromMenuSet Me
        
        DestroyMenu m_hMenu
        m_hMenu = 0&
        
    End If
    
    Set m_Col = Nothing
    Set m_Col = New Collection

End Sub

'''' Public, flexible Load and Save properties with optional prefix
'''' for saving to other control's property bags

'' Load

Public Sub LoadProperties(PropBag As PropertyBag, Optional Prefix As String)
    Dim sPrefix As String, _
        i As Long, _
        j As Long, _
        objItem As MenuItem
    
    Dim fBool As Boolean
    
    If (Prefix <> "") Then sPrefix = Prefix + "."
        
    m_ItemId = PropBag.ReadProperty(sPrefix + "ItemId", m_ItemId)
    m_RightToLeft = PropBag.ReadProperty(sPrefix + "RightToLeft", m_RightToLeft)
    
    m_MaxHeight = PropBag.ReadProperty(sPrefix + "MaxHeight", m_MaxHeight)
    m_AutoDismiss = PropBag.ReadProperty(sPrefix + "AutoDismiss", m_AutoDismiss)
    
    m_LargeIcons = PropBag.ReadProperty(sPrefix + "LargeIcons", m_LargeIcons)
    m_Key = PropBag.ReadProperty(sPrefix + "Key", m_Key)
    
    Set m_Font = PropBag.ReadProperty(sPrefix + "Font", m_Font)
        
    i = PropBag.ReadProperty(sPrefix + "Count", i)
    
    Me.Clear
    For j = 1 To i
        Set objItem = New MenuItem
        Set objItem.Parent = Me
        
        objItem.Index = (j - 1&)
        objItem.LoadProperties PropBag, sPrefix + "Item" & j
    
        m_Col.Add objItem, objItem.Key
        Set objItem = Nothing
            
    Next j
    
    fBool = PropBag.ReadProperty(sPrefix + "HasSidebar", fBool)
    
    If (fBool = True) Then
        Set m_Sidebar = New Sidebar
        Set m_Sidebar.Parent = Me
        
        m_Sidebar.LoadProperties PropBag, sPrefix + "Sidebar"
    End If
    
    Set m_Accelerators = New Accelerators
    UpdateTables Me
    
End Sub
 
'' Save
 
Public Sub SaveProperties(PropBag As PropertyBag, Optional Prefix As String)
    Dim sPrefix As String, _
        i As Long, _
        j As Long, _
        objItem As MenuItem
    
    Dim fBool As Boolean
    
    If (Prefix <> "") Then sPrefix = Prefix + "."

    PropBag.WriteProperty sPrefix + "ItemId", m_ItemId
    PropBag.WriteProperty sPrefix + "RightToLeft", m_RightToLeft
    
    PropBag.WriteProperty sPrefix + "AutoDismiss", m_AutoDismiss
    PropBag.WriteProperty sPrefix + "MaxHeight", m_MaxHeight
    
    PropBag.WriteProperty sPrefix + "LargeIcons", m_LargeIcons
    PropBag.WriteProperty sPrefix + "Key", m_Key
    
    PropBag.WriteProperty sPrefix + "Font", m_Font
    
    i = m_Col.Count
            
    PropBag.WriteProperty sPrefix + "Count", i
    
    For j = 1 To i
        Set objItem = m_Col(j)
        
        objItem.SaveProperties PropBag, sPrefix + "Item" & j
    Next j
    
    Set objItem = Nothing

    If Not m_Sidebar Is Nothing Then
        fBool = True
    End If
    
    PropBag.WriteProperty sPrefix + "HasSidebar", fBool
    
    If (fBool = True) Then
        m_Sidebar.SaveProperties PropBag, sPrefix + "Sidebar"
    End If
    
End Sub
   
'' Friend properties for internal management routines

Friend Property Get Col() As Collection
    Set Col = m_Col
End Property
        
Friend Property Get WindowListCol() As Collection
    Set WindowListCol = g_WindowList
End Property



'' This method is called by MenuWndProc() when a menu item is selected
'' that exists anywhere in the global collection for a thread.
'' each ExecCmd() call forwards the message to the object identified by the
'' TopMostParent property.  If this is the TopMostParent, then an event is raised, instead.

Friend Function ExecCmd(ByVal wParam As Long, ByVal uMsg As Long) As Long
    '' This is the SubMenu version
        
    On Error Resume Next
        
    Dim objItem As Object, _
        fRight As Boolean
    
    Select Case uMsg
    
        Case WM_COMMAND, WM_MENURBUTTONUP
        
            If (uMsg = WM_MENURBUTTONUP) Then fRight = True
            Set objItem = g_MenuCol.SearchRecursive(Me, wParam)
            
            If objItem Is Nothing Then Exit Function
            
            If TypeOf objItem Is Sidebar Then
                
                If (fRight = True) Then
                    RaiseEvent SidebarRightClick
                Else
                    RaiseEvent SidebarClick
                End If
            
            ElseIf TypeOf objItem Is MenuItem Then
                
                If (fRight = True) Then
                    RaiseEvent ItemRightClick(objItem)
                Else
                    RaiseEvent UserCommand(objItem)
                End If
            Else
                Set objItem = Nothing
                Exit Function
            End If
            
        Case WM_INITMENUPOPUP
            
            If (wParam = m_hMenu) Then
                RaiseEvent Initialize
            Else
                Set objItem = g_MenuCol.MenuByHandle(wParam)
                If Not objItem Is Nothing Then
                    If IsParent(Me, objItem) Then
                        RaiseEvent SubmenuInitialize(objItem)
                    End If
                    
                    Set objItem = Nothing
                End If
            End If
        
    End Select
    
    If Not m_Parent Is Nothing Then
        SendCommand TopMostParent, wParam, uMsg
    End If

    Set objItem = Nothing
    
End Function


'' Clear the RadioGroup selections before selecting a new
'' item (called by child items whose properties change)

Friend Sub ClearRadioGroup()
    Dim objItem As MenuItem
    
    For Each objItem In Me
        If (objItem.RadioGroup = True) Then
            objItem.Checked = False
        End If
    Next objItem
    
End Sub
   
'' Private Read/Write Properties (for persistence)

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    SaveProperties PropBag
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    LoadProperties PropBag
End Sub

Private Sub InsertItem(ByVal Item As MenuItem, InsertBefore)

End Sub

'' Private Subroutines called when an instance is created or destroyed

Private Sub Class_Initialize()
    
    Set m_Col = New Collection
    Set m_Accelerators = New Accelerators
    Set m_Accelerators.Menu = Me
        
    m_ItemId = GetNewCtrlId
    
    Set m_Sidebar = New Sidebar
    Set m_Sidebar.Parent = Me

End Sub

Private Sub Class_Terminate()
    
    If (m_hMenu <> 0&) Then Destroy
    
    Set m_Sidebar = Nothing
    Set m_Col = Nothing
    
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






