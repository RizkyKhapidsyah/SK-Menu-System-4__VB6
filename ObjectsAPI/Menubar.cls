VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Menubar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"UserDefinedCollection"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
''' GUInerd Standard Menu System
''' Version 4.1

''' Objects/API Dll

''' Menubar Class Object


''' ****************************************************************************
''' *** Check the end of the file for copyright and distribution information ***
''' ****************************************************************************

Option Explicit

Private m_Counter As Long

''' Private variables that represent public or read-only properties

Private m_Connected As Boolean

'' Identification

Private m_ItemId As Long

Private m_Key As String

'' hMenu

Private m_hMenu As Long

'' Window-related

Private m_OldWndProc As Long

Private m_OldWndMenu As Long

Private m_hWnd As Long

'' Information about this menu stored in Win32 standard format

Private m_ItemInfo As MENUITEMINFO

'' Collection to store child items

Private m_Col As Collection

'' Accelerator collection

Private m_Accelerators As Accelerators

'' Window menu property (for MDI forms)

Private m_WindowList As Submenu

'' Default Font property

Private m_Font As StdFont

' Right To Left Reading Order

Private m_RightToLeft As Boolean

' In Win98/Win ME, we need to test the WM_NCMOUSEMOVE statement
' this stores any previous result

Private m_LastOver As MenuItem

'' Max image widths that are recorded on refresh
Private m_miwLeft As Long

Private m_miwRight As Long

'' Events

Public Event Initialize()

Public Event UserCommand(ByVal Item As MenuItem)

Public Event ItemRightClick(ByVal Item As MenuItem)

Public Event SubmenuInitialize(ByVal Submenu As Submenu)

' Menu types.

Public Enum EnhancedMenuTypeConstants
    emPopupMenu = &H1&
    emBarMenu = &H2&
End Enum


''' Window-related properties
'''

' Owner Window hMenu

Public Property Get hWnd() As Long
    hWnd = m_hWnd
End Property

' Old window procedure

Public Property Get OldWndProc() As Long
    OldWndProc = m_OldWndProc
End Property

' Old window menu

Public Property Get OldWndMenu() As Long
    OldWndMenu = m_OldWndMenu
End Property

'''

''' Menu-related properties
'''

' Menu handle

Public Property Get hMenu() As Long
    hMenu = m_hMenu
End Property

''' The Subclass method may need to set this variable externally

Friend Property Let hMenu(ByVal vData As Long)
    On Error Resume Next
    
    If m_hMenu <> 0& Then Destroy
    
    m_hMenu = vData
    
    If g_MenuCol Is Nothing Then Set g_MenuCol = New Menus
    
    If m_hMenu <> 0& Then
        If Not g_MenuCol("_H" + Hex(m_hMenu)) Is Nothing Then
            g_MenuCol.Remove "_H" + Hex(m_hMenu)
        End If
        
        g_MenuCol.Add Me, "_H" + Hex(m_hMenu)
    End If
    
End Property

Public Property Get Key() As String
    Key = m_Key
End Property

Public Property Let Key(ByVal vData As String)
    m_Key = vData
End Property

Public Property Get Connected() As Boolean
    Connected = m_Connected
End Property

Public Property Let Connected(ByVal vData As Boolean)
    m_Connected = vData
End Property

''' The Menu ID is assigned when the menu object is created, but
''' re-assigned when a menu is added to the collection.
''' The GetNewCtrlId function generates a CtrlID at the top most level
''' of the menu heirarchy.

''' The menu ID for a menu bar toplevel menu is arbitrary, it is not
''' physically assigned in Windows.

Public Property Get ItemId() As Long
    ItemId = m_ItemId
End Property

Friend Property Let ItemId(ByVal vData As Long)
    m_ItemId = vData
End Property

'' Keyboard accelerator collection property

Public Property Get Accelerators() As Accelerators
    On Error Resume Next
    
    Set Accelerators = m_Accelerators

End Property

'''

' Font property

Public Property Get Font() As StdFont

    Set Font = m_Font
    
End Property

Public Property Set Font(IFont As StdFont)

    Set m_Font = IFont
    
    If m_hWnd <> 0& Then
        DrawMenuBar m_hWnd
    End If
    
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

' Function to set the MENUITEMINFO structure
' from the owning Item's MENUITEMINFO.
' The item's info is valid and updated
' every time the parent item's SetCtrlState or GetCtrlState functions are called.
' This is to ensure that the Submenu has the proper data variables to pass to
' the menu handler when creating new sub items.

Friend Sub SetItemInfo(lpInfo As MENUITEMINFO)
    
    CopyMemory m_ItemInfo, lpInfo, Len(m_ItemInfo)
    
End Sub

'' This method is called by MenuWndProc() when a menu item is selected
'' that exists anywhere in the global collection for a thread.
'' each ExecCmd() call forwards the message to the object identified by the
'' TopMostParent property.  Menubar is always TopMost.

Friend Function ExecCmd(ByVal wParam As Long, ByVal uMsg As Long) As Long
    '' This is the Menubar version
        
    On Error Resume Next
        
    Dim objItem As Object, _
        fRight As Boolean
        
                
    Select Case uMsg
    
        Case WM_COMMAND, WM_MENURBUTTONUP
        
            If (uMsg = WM_MENURBUTTONUP) Then fRight = True
            
            Set objItem = g_MenuCol.SearchRecursive(Me, wParam)
            
            If objItem Is Nothing Then Exit Function
            
            If Not TypeOf objItem Is MenuItem Then Exit Function
            
            If (fRight = True) Then
                RaiseEvent ItemRightClick(objItem)
            Else
                RaiseEvent UserCommand(objItem)
            End If
            
            Set objItem = Nothing
    
        Case WM_INITMENU
            If (wParam = m_hMenu) Then
                RaiseEvent Initialize
            End If
            
        Case WM_INITMENUPOPUP
            Set objItem = g_MenuCol.MenuByHandle(wParam)
            
            If Not objItem Is Nothing Then
                
                If IsParent(Me, objItem) Then
                    RaiseEvent SubmenuInitialize(objItem)
                End If
                
                Set objItem = Nothing
            End If
    
    End Select
    
End Function



'' This property is constant in an Menubar

Public Property Get TopMostParent() As Object
    Set TopMostParent = Me
End Property

'''''' Public properties and methods relating to the menu item collection

'' Add Method

Public Function Add(ByVal sCaption As String, Optional Picture As StdPicture, Optional ByVal sKey As String, Optional InsertAfter) As MenuItem
    On Error Resume Next
    
    Dim lpInfo As MENUITEMINFO, _
        cbFlags As Long, _
        varStr As String
    
    Dim objNewItem As MenuItem, _
        objBefore As MenuItem
        
    If m_Col Is Nothing Then Set m_Col = New Collection
    
    lpInfo.cbSize = g_InfoSize
        
    If m_ItemId = 0& Then Exit Function
    
    ''' Create this menu as necessary...
    
    If m_hMenu = 0& Then Create
                           
    ''' Set the parent of the new item to this instance.
    
    Set objNewItem = New MenuItem
    Set objNewItem.Parent = Me
    Set objNewItem.Accelerator.Table = m_Accelerators
    
    ''' If there is no key, it is the hexidecimal value of the control id.
    
    If sKey = "" Then
        objNewItem.Key = "_H" + Hex(objNewItem.ItemId)
    Else
        objNewItem.Key = sKey
    End If
    
    ''' Menu items are owner-drawn.
    
    cbFlags = MF_OWNERDRAW + MF_BYCOMMAND
    
    ''' Use the old AppendMenu to quickly add a menu item.
    
    If (IsMissing(InsertAfter) = False) Then
        If VarType(InsertAfter) = vbString Then
            Set objBefore = m_Col(InsertAfter)
        ElseIf IsObject(InsertAfter) Then
            Set objBefore = InsertAfter
        Else
            Set objBefore = Me.Search(InsertAfter)
        End If
    End If
    
    If Not objBefore Is Nothing Then
        InsertMenu m_hMenu, objBefore.ItemId, cbFlags, CLng(objNewItem.ItemId), 0&
        
        m_Col.Add objNewItem, objNewItem.Key, , objBefore.Index + 1
        ReindexMenu Me
        
    Else
        objNewItem.Index = m_Col.Count
        
        AppendMenu m_hMenu, cbFlags, objNewItem.ItemId, 0&
        ''' Add the item to the collection
        
        m_Col.Add objNewItem, objNewItem.Key
    End If
    
    
    ''' Now set the item's properties so that they are picked up
    ''' by the physical control...
    
    objNewItem.RightToLeft = m_RightToLeft
    objNewItem.Visual.TextAlign = taCenter
    
    objNewItem.Caption = sCaption
    
    If Not Picture Is Nothing Then
        Set objNewItem.Picture = Picture
    End If
        
    '' If the global drawing Style is for OfficeXP, pre-configure
    '' menu items (this is to allow flexibility to remain within the item
    '' as much as possible.  These configurations are suggestions, they can
    '' change, with dramatic results of done tastefully and carefully!)
    
    If (g_MenuCol.MenuDrawStyle = mdsOfficeXP) Then
        objNewItem.Visual.SelectionStyle = (mssFlat + mssColor)
    End If
        
    objNewItem.Unfreeze
    objNewItem.SetCtrlState
        
    ' Since this is a menu bar object, we redraw the menu bar to reflect the
    ' changes.
    
    If (m_hWnd <> 0&) And (Connected = False) Then
        DrawMenuBar m_hWnd
    End If
    
    Set Add = objNewItem
    Set objNewItem = Nothing
    
End Function

'' Get an item

Public Property Get Item(varIndexKey) As MenuItem
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
  On Error Resume Next
  Set Item = m_Col(varIndexKey)
End Property

Public Property Get Submenu(varIndexKey) As Submenu
  On Error Resume Next
  Set Submenu = Item(varIndexKey).Submenu
End Property



'' Get the item count

Public Property Get Count() As Long
    Count = m_Col.Count
End Property

'' Destroy and remove a menu item

Public Sub Remove(varIndexKey As Variant)
    On Error Resume Next
    
    If Not m_Col(varIndexKey) Is Nothing Then
        m_Col(varIndexKey).Destroy
        m_Col.Remove varIndexKey
    End If
    
    ReindexMenu Me
End Sub

'' Destroy all sub menus and create a new collection

''' Destroy the sub menus first, or we will generate an exception.
''' You must destroy the menus before terminating the class.

Public Sub Clear()
    On Error Resume Next
    
    Dim objItem As MenuItem, _
        i As Long, _
        j As Long
    
    For Each objItem In m_Col
        objItem.Submenu.Destroy
        objItem.Destroy
    Next objItem
    
    Set m_Col = Nothing
    Set m_Col = New Collection
    
End Sub

'' NewEnum function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_Col.[_NewEnum]
End Function

Friend Property Get Col() As Collection
    Set Col = m_Col
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

Public Property Get Child() As MenuItem
    If (m_Col.Count <> 0&) Then
        Set Child = m_Col(1)
    Else
        Set Child = Nothing
    End If
End Property


''' Properties and methods used in connecting to MDI applications

'' Connection property (applies to usage in MDI applications)

'' WindowList property (for MDI forms)

Public Property Get WindowList() As Submenu
    Set WindowList = m_WindowList
End Property

Public Sub SetWindowList(Menu)
    Dim dwHandle As Long
    Dim objItem As MenuItem

    If IsObject(Menu) Then
    
        If TypeOf Menu Is Submenu Then
            If Menu.TopMostParent Is Me Then
                Set m_WindowList = Menu
            End If
            
        End If
        
    Else
        Set m_WindowList = m_Col(Menu).Submenu
    End If
                 
    If m_WindowList Is Nothing Then Exit Sub
    
    dwHandle = m_WindowList.hMenu
     
    For Each objItem In m_Col
    
        If objItem.Submenu.hMenu <> dwHandle Then
            objItem.Submenu.WindowList = False
        Else
            objItem.Submenu.WindowList = True
        End If
        
    Next objItem

    If m_Connected = True Then
    
        g_WindowListHandle = dwHandle
    
'        RefreshChildWindows m_Connection.hWndFrame
        
        g_IsSetting = True
        
'        SendMessage m_Connection.hWndClient, WM_MDISETMENU, 0&, dwHandle
'        SendMessage m_Connection.hWndClient, WM_MDIREFRESHMENU, 0&, 0&
        
        g_IsSetting = False
    
'        DrawMenuBar m_Connection.hWndFrame
        
    End If

End Sub

'' Search method passes values on to the global SearchRecursive routine

Public Function Search(varIndexKey As Variant) As MenuItem

    ''' See SearchRecursive method in m_Menubar
    Set Search = g_MenuCol.SearchRecursive(Me, varIndexKey)
    
End Function


'' Method to cause this object to take over the drawing responsibilities
'' for an existing menu.  This object becomes, then, an 'interface' to that
'' menu, and the internal m_hMenu identifier is changed to the menu that is
'' being subclassed.


'' Method to cause this object to take over the drawing responsibilities
'' for an existing menu.  This object becomes, then, an 'interface' to that
'' menu, and the internal m_hMenu identifier is changed to the menu that is
'' being subclassed.

Public Sub Subclass(ByVal hWnd As Long, Optional ByVal fTagOriginalIDs As Boolean, Optional ByVal fTranslate As TranslateItemDataConstants)
    On Error Resume Next
            
    If (m_hWnd <> 0&) Then Exit Sub
    
    If (m_hMenu <> 0&) Then
        Destroy
    End If
    
    m_hWnd = hWnd
    m_hMenu = GetMenu(m_hWnd)
    
    CopyMenubar_API m_hMenu, Me, True, fTagOriginalIDs, fTranslate
    SetWindowHandle m_hWnd
    
    AddToMenuSet Me
    
End Sub

Public Sub Redraw()
    If m_hWnd <> 0& Then
        DrawMenuBar m_hWnd
    End If
End Sub

Public Sub CopyFromWindow(ByVal hWnd As Long, Optional ByVal fTagOriginalIDs As Boolean, Optional ByVal fTranslate As TranslateItemDataConstants)
    
    Dim cpMenu As Long
    
    cpMenu = GetMenu(hWnd)
    CopyMenu cpMenu, fTagOriginalIDs, fTranslate

End Sub

'' Method to copy the contents of another menu into this one

Public Sub CopyMenu(ByVal hMenu As Long, Optional ByVal fTagOriginalIDs As Boolean, Optional ByVal fTranslate As TranslateItemDataConstants)
    
    If (m_hMenu <> 0&) Then
        Destroy
    End If
    
    m_hMenu = CreateMenu
    AddToMenuSet Me
    
    CopyMenubar_API hMenu, Me, False, fTagOriginalIDs, fTranslate
    
    If m_hWnd <> 0& Then
        SetMenu m_hWnd, m_hMenu
        DrawMenuBar m_hWnd
    End If

End Sub


'' Release the window handle from the subclassing procedure and
'' remove associations between this instance and the assigned window

Public Sub ReleaseWindowHandle()

    On Error Resume Next
        
    If m_hWnd <> 0& Then
        
        If Connected = False Then
            RemoveHandle m_hWnd
        End If
        
        m_hWnd = 0&
    End If
    
End Sub

    ' We only want top level menus to hook to windows.
    ' popup menus can hook through the Popup subroutine,
    ' or directly with a call to TrackPopupMenu.
    
Public Sub SetWindowHandle(ByVal hWnd As Long, Optional ByVal CopyOldMenu As Boolean = False)
    On Error Resume Next
    Dim sKey As String
    
        ' cannot assign without first releasing.
        
    If m_hWnd <> 0& Then Exit Sub
    
    m_hWnd = hWnd
    
    If (m_Connected = False) Then
    
        m_OldWndMenu = GetMenu(m_hWnd)
        
        If m_OldWndMenu <> m_hMenu Then
            SetMenu m_hWnd, m_hMenu
                
            If CopyOldMenu = True Then
                CopyMenu_API m_OldWndMenu, Me, False
            End If
        Else
            m_OldWndMenu = 0&
        End If
                
    
        If GetOldWndProc(m_hWnd) = 0& Then
            AddHandle m_hWnd, AddressOf MenuWndProc, m_hMenu
        End If
        
        If (m_Col.Count <> 0&) Then
            RecreateMenu Me
        End If
    
    End If
        
    DrawMenuBar m_hWnd

End Sub


Public Sub SetWindowHandleMDI(ByVal hMDIWnd As Long, Optional ByVal CopyOldMenu As Boolean = False)
    On Error Resume Next
    Dim sKey As String
    
        ' cannot assign without first releasing.
        
    If m_hWnd <> 0& Then Exit Sub
    
    m_hWnd = hWnd
    
    If (m_Connected = False) Then
    
        m_OldWndMenu = GetMenu(m_hWnd)
        
        If m_OldWndMenu <> m_hMenu Then
            SetMenu m_hWnd, m_hMenu
                
            If CopyOldMenu = True Then
                CopyMenu_API m_OldWndMenu, Me, False
            End If
        Else
            m_OldWndMenu = 0&
        End If
                
    
        If GetOldWndProc(m_hWnd) = 0& Then
            AddHandle m_hWnd, AddressOf MenuWndProc, m_hMenu
        End If
        
        If (m_Col.Count <> 0&) Then
            RecreateMenu Me
        End If
    
    End If
        
    DrawMenuBar m_hWnd

End Sub


'' Clear the RadioGroup selections before selecting a new
'' item

Friend Sub ClearRadioGroup()
    Dim objItem As MenuItem
    
    For Each objItem In Me
        If (objItem.RadioGroup = True) Then
            objItem.Checked = False
        End If
    Next objItem
    
End Sub
   
'' Destroy is called internally by the Class_Terminate and Subclass Methods
'' It also can be called externally (e.g. when an error has occured, or to quickly
'' delete the menu without having to unload the instance.)

Public Sub Destroy()
    Dim objItem As MenuItem
    
        On Error Resume Next
    
    If (m_hWnd <> 0&) Then
        ReleaseWindowHandle
    End If
                        
    For Each objItem In m_Col
        objItem.Submenu.Destroy
        objItem.Destroy
    Next objItem
    
    If m_hMenu <> 0& Then
        RemoveFromMenuSet Me
        
        DestroyMenu m_hMenu
        m_hMenu = 0&
    End If
                        
    
    Set m_Col = Nothing
    Set m_Col = New Collection
    
End Sub

Public Sub Create()
    On Error Resume Next
       
    If m_hMenu <> 0& Then Exit Sub
    
    m_hMenu = CreateMenu
    
    AddToMenuSet Me
    
    If (m_hWnd <> 0&) And (m_Connected = False) Then
        SetMenu m_hWnd, m_hMenu
        DrawMenuBar m_hWnd
    End If
    
    If m_Col Is Nothing Then Set m_Col = New Collection
End Sub

Friend Property Get LastOver() As MenuItem
    Set LastOver = m_LastOver
End Property

Friend Property Set LastOver(ByVal vData As MenuItem)
    Set m_LastOver = vData
End Property



Public Sub LoadProperties(PropBag As PropertyBag, Optional Prefix As String)
    Dim sPrefix As String, _
        i As Long, _
        j As Long, _
        objItem As MenuItem
    
    If (Prefix <> "") Then sPrefix = Prefix + "."
        
    m_ItemId = PropBag.ReadProperty(sPrefix + "ItemId", m_ItemId)
    m_RightToLeft = PropBag.ReadProperty(sPrefix + "RightToLeft", m_RightToLeft)
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
    
    Set m_Accelerators = New Accelerators
    UpdateTables Me
    
End Sub
 
Public Sub SaveProperties(PropBag As PropertyBag, Optional Prefix As String)
    Dim sPrefix As String, _
        i As Long, _
        j As Long, _
        objItem As MenuItem
    
    If (Prefix <> "") Then sPrefix = Prefix + "."

    PropBag.WriteProperty sPrefix + "ItemId", m_ItemId
    PropBag.WriteProperty sPrefix + "RightToLeft", m_RightToLeft
    PropBag.WriteProperty sPrefix + "Key", m_Key
    
    PropBag.WriteProperty sPrefix + "Font", m_Font
    
    i = m_Col.Count
            
    PropBag.WriteProperty sPrefix + "Count", i
    
    For j = 1 To i
        Set objItem = m_Col(j)
        
        objItem.SaveProperties PropBag, sPrefix + "Item" & j
    Next j
    
    Set objItem = Nothing

End Sub
   
'' Private Read/Write Properties

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    SaveProperties PropBag
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    LoadProperties PropBag
End Sub

'' Private Subroutines called when an instance is created or destroyed

Private Sub Class_Initialize()
    Set m_Col = New Collection
        
    m_ItemInfo.cbSize = g_InfoSize
    
    ' Create a new accelerator collection

    Set m_Accelerators = New Accelerators
    Set m_Accelerators.Menu = Me
                
    ' Get new global identifier.  Menu IDs are unique throughout a thread
        
    m_ItemId = GetNewCtrlId
    
    '' Create a menu handle, add it to the menu array.
            
    Create
        
End Sub

Private Sub Class_Terminate()
    
    Set m_LastOver = Nothing
    
  ' Automatically destroy the whole thing on terminate.
  '' (See notes in 'Destroy' method, above)
    
    ReleaseWindowHandle
    Destroy
    
  
  '' It is standard to clear collections on termination
    
    Set m_Col = Nothing
    Set m_Accelerators = Nothing
    
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






Public Property Get Counter() As Long
    Counter = m_Counter
End Property

Friend Property Let Counter(ByVal vData As Long)
    m_Counter = vData
End Property
