VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Menus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' GUInerd Standard Menu System
''' Version 4.1

''' Objects/API Dll

''' Menus Collection Object


''' ****************************************************************************
''' *** Check the end of the file for copyright and distribution information ***
''' ****************************************************************************

Option Explicit

Public Enum TranslateItemDataConstants
    tdNone = &H0
    tdBinary = &H1
    tdString = &H2
End Enum

Public Enum MenuTypeConstants
    mtcMenubar = 0&
    mtcPopup = 1&
    mtcSysmenu = 2&
    mtcDesigner = 3&
End Enum

Public Enum MenuDrawStyleBits
    mdsOffice = &H0
    mdsOfficeXP = &H2
End Enum

Public Enum AcceleratorStyleBits
    sbNormalKeys = &H0&
    sbExtendedKeys = &H1&
End Enum

Private m_GlobalNoUnicode As Boolean

Private m_PrefixLib As Accelerators

Private m_AcceleratorStyle As AcceleratorStyleBits

Private m_Active As Boolean

Private m_MenuDrawStyle As MenuDrawStyleBits

Private m_Col As Collection

Public Property Get MenuDrawStyle() As MenuDrawStyleBits
    MenuDrawStyle = m_MenuDrawStyle
End Property

Public Property Let MenuDrawStyle(ByVal vData As MenuDrawStyleBits)
    Dim lpMenu As Object
    
    m_MenuDrawStyle = vData
    
    For Each lpMenu In m_Col
    
        RecreateMenu lpMenu
    
    Next lpMenu
    
End Property

Public Property Get GlobalNoUnicode() As Boolean
    GlobalNoUnicode = m_GlobalNoUnicode
End Property

Public Property Let GlobalNoUnicode(ByVal vData As Boolean)
    m_GlobalNoUnicode = vData
    CheckUnicode
End Property

Public Property Get PrefixLib() As Accelerators
    Set PrefixLib = m_PrefixLib
End Property

Public Property Get AcceleratorStyle() As AcceleratorStyleBits
    AcceleratorStyle = m_AcceleratorStyle
End Property

Public Property Let AcceleratorStyle(ByVal vData As AcceleratorStyleBits)
    m_AcceleratorStyle = vData
End Property

Public Property Get Active() As Boolean
    Active = m_Active
End Property

Friend Property Let Active(ByVal vData As Boolean)
    m_Active = vData
End Property

Public Sub AddDesigner(Designer As Object, Optional ByVal Key As String)
    On Error Resume Next
    Dim varObj As Object, _
        fIsDesigner As Boolean
        
    fIsDesigner = Designer.IsMenuDesigner
        
    If (TypeOf Designer Is Menubar) Or (TypeOf Designer Is Submenu) Or _
        (TypeOf Designer Is SystemMenu) Or (fIsDesigner = False) Then
        
        Err.Raise 13&, "StdMenuAPI.Menus"
        Exit Sub
    End If
    
    If Key <> "" Then
        Set varObj = m_Col(Key)
        If Not varObj Is Nothing Then
            m_Col.Remove Key
        End If
    End If
        
    m_Col.Add Designer, Key
    
End Sub

Public Sub Add(Menu As Object, Optional ByVal Key As String)
    On Error Resume Next
    Dim varObj As Object
    
    If (Me Is g_MenuCol) Then
        m_Active = True
    End If
    
    If (Not TypeOf Menu Is Menubar) And (Not TypeOf Menu Is Submenu) And _
        (Not TypeOf Menu Is SystemMenu) Then
        
        Err.Raise 13&, "StdMenuAPI.Menus"
        Exit Sub
    End If
    
    If Key <> "" Then
        Set varObj = m_Col(Key)
        If Not varObj Is Nothing Then
            m_Col.Remove Key
        End If
    End If
        
    Menu.Key = Key
    m_Col.Add Menu, Key
    
End Sub

Public Sub Remove(varIndexKey)
    On Error Resume Next
    If (m_Col(varIndexKey) Is Nothing) Then
        Err.Raise 9&, "StdMenuAPI.Menus"
    Else
        m_Col.Remove varIndexKey
    End If
End Sub

Public Property Get Item(varIndexKey) As Object
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
    On Error Resume Next
    
    Dim varObj As Object
    
    Set varObj = m_Col(varIndexKey)
    If (Not varObj Is Nothing) Then
        Set Item = varObj
    Else
        Err.Raise 9&, "StdMenuAPI.Menus"
    End If

End Property

Public Function MenuByHandle(ByVal hMenu As Long, Optional ByVal lParam As Long) As Object
    Dim objMenu As Object, _
        i As Long
    
    For Each objMenu In m_Col
        i = objMenu.hMenu
        
        If i = hMenu Then
            If ((lParam And &HFFFF0000) <> 0&) Then
                If TypeOf objMenu Is SystemMenu Then
                    Set MenuByHandle = objMenu
                    Set objMenu = Nothing
                    Exit Function
                End If
            Else
                Set MenuByHandle = objMenu
                Set objMenu = Nothing
                Exit Function
            End If
        End If
    Next objMenu
    
End Function

Public Property Get HasMenu(ByVal hMenu As Long) As Boolean
    Dim objMenu As Object, _
        i As Long
    
    For Each objMenu In m_Col
        i = objMenu.hMenu
        If i = hMenu Then
            HasMenu = True
            Set objMenu = Nothing
            Exit Property
        End If
    Next objMenu
    
End Property

Public Property Get ItemType(varIndexKey) As MenuTypeConstants
    On Error Resume Next
    
    Dim varObj As Object
    
    Set varObj = m_Col(varIndexKey)
    If (Not varObj Is Nothing) Then
        If TypeOf varObj Is Menubar Then
            ItemType = mtcMenubar
        ElseIf TypeOf varObj Is Submenu Then
            ItemType = mtcPopup
        ElseIf TypeOf varObj Is SystemMenu Then
            ItemType = mtcPopup
        Else
            ItemType = mtcDesigner
        End If
    Else
        Err.Raise 9&, "StdMenuAPI.Menus"
    End If

End Property

Public Property Get Submenu(varIndexKey) As Submenu
    On Error Resume Next
    
    Dim varObj As Object
    
    Set varObj = m_Col(varIndexKey)
    If (Not varObj Is Nothing) And (TypeOf varObj Is Submenu) Then
        Set Submenu = varObj
        Set varObj = Nothing
    ElseIf varObj Is Nothing Then
        Err.Raise 9&, "StdMenuAPI.Menus"
    Else
        Err.Raise 13&, "StdMenuAPI.Menus", "This is an internal error.  Please contact the author."
    End If

End Property

Public Property Get Menubar(varIndexKey) As Menubar
    On Error Resume Next
    
    Dim varObj As Object
    
    Set varObj = m_Col(varIndexKey)
    If (Not varObj Is Nothing) And (TypeOf varObj Is Menubar) Then
        Set Menubar = varObj
        Set varObj = Nothing
    ElseIf varObj Is Nothing Then
        Err.Raise 9&, "StdMenuAPI.Menus"
    Else
        Err.Raise 13&, "StdMenuAPI.Menus", "This is an internal error.  Please contact the author."
    End If

End Property

Public Property Get SystemMenu(varIndexKey) As SystemMenu
    On Error Resume Next
    
    Dim varObj As Object
    
    Set varObj = m_Col(varIndexKey)
    If (Not varObj Is Nothing) And (TypeOf varObj Is SystemMenu) Then
        Set SystemMenu = varObj
        Set varObj = Nothing
    ElseIf varObj Is Nothing Then
        Err.Raise 9&, "StdMenuAPI.Menus"
    Else
        Err.Raise 13&, "StdMenuAPI.Menus", "This is an internal error.  Please contact the author."
    End If

End Property

Public Property Get Designer(varIndexKey) As Object
    On Error Resume Next
    
    Dim varObj As Object
    
    Set varObj = m_Col(varIndexKey)
    If (Not varObj Is Nothing) And (Not TypeOf varObj Is SystemMenu) And _
        (Not TypeOf varObj Is Menubar) And (Not TypeOf varObj Is Submenu) Then
        
        Set Designer = varObj
        Set varObj = Nothing
    ElseIf varObj Is Nothing Then
        Err.Raise 9&, "StdMenuAPI.Menus"
    Else
        Err.Raise 13&, "StdMenuAPI.Menus", "This is an internal error.  Please contact the author."
    End If

End Property

Public Property Get GlobalFont() As StdFont
    Set GlobalFont = g_SysMenuFont
End Property

Public Property Set GlobalFont(ByVal vData As StdFont)
    Set g_SysMenuFont = vData
End Property

Public Property Get Count() As Long
    On Error Resume Next
    Count = m_Col.Count
End Property

Public Property Get MenuCount() As Long
    Dim varObj As Object, _
        i As Long
        
    For Each varObj In m_Col
        If (TypeOf varObj Is Submenu) Or _
            (TypeOf varObj Is Menubar) Or _
            (TypeOf varObj Is SystemMenu) Then
            
            i = i + 1
        End If
    Next varObj
    
    MenuCount = i
End Property

Public Property Get MenubarCount() As Long
    Dim varObj As Object, _
        i As Long
        
    For Each varObj In m_Col
        If (TypeOf varObj Is Menubar) Then
            
            i = i + 1
        End If
    Next varObj
    
    MenubarCount = i
End Property

Public Property Get SubmenuCount() As Long
    Dim varObj As Object, _
        i As Long
        
    For Each varObj In m_Col
        If (TypeOf varObj Is Submenu) Then
            
            i = i + 1
        End If
    Next varObj
    
    SubmenuCount = i
End Property

Public Property Get SystemMenuCount() As Long
    Dim varObj As Object, _
        i As Long
        
    For Each varObj In m_Col
        If (TypeOf varObj Is SystemMenu) Then
            
            i = i + 1
        End If
    Next varObj
    
    SystemMenuCount = i
End Property

Public Property Get DesignerCount() As Long
    Dim varObj As Object, _
        i As Long
    
    Dim fBool As Boolean
    
    On Error Resume Next
    
    For Each varObj In m_Col
        fBool = varObj.IsDesigner
        
        If (fBool = True) Then
            i = i + 1
            fBool = False
        End If
    Next varObj
    
    DesignerCount = i
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_Col.[_NewEnum]
End Function

Public Sub Clear()
    Set m_Col = Nothing
    Set m_Col = New Collection
End Sub

' Find a menu given an item ID.

' Because it is possible to have more than one menu on a window or as part of this thread,
' we will assign all menu identifiers incrementally from the first menu created
' on the first window destination to the last menu item in the last Submenu in the last
' window, using a global identifier that is incremented by every created instance of
' the Menubar object.

Public Function FindMenuItem(ByVal ItemId As Long, Optional ByVal hWnd As Long) As Object
        
    On Error Resume Next

    ' Calling late-bound objects in callback processes is not a good idea
    ' and is somewhat unsuccessful.  Call each according to its type.

    Dim objMenu As Object
    Dim objItem As Object
    Dim i As Long
    
    ' Iterate the collection, return on success.
    
    For i = 1 To m_Col.Count
        
        Set objMenu = Nothing
        Set objMenu = m_Col(i)
        
        Set objItem = SearchRecursive(objMenu, ItemId, hWnd)
        
        If Not objItem Is Nothing Then
            Set FindMenuItem = objItem
            Exit Function
        End If
    
    Next i
                
    Set FindMenuItem = Nothing
    
End Function

    ' The search function will read the whole data collection.
    ' If it is textual, then it will search every menu and every
    ' Items of every menu and so forth looking for a matching key.
    ' If it's an integer value, then we assume, that it refers directly to the
    ' ItemId assigned by the GetCtrlID function in Menubar
        
Public Function SearchRecursive(ByVal objMenu As Object, ByVal ItemId As Long, Optional ByVal hWnd As Long) As Object
    Dim varObj As MenuItem, _
        varBar As Sidebar
        
    Dim varFind As Object, _
        varSys As SystemMenu
        
    On Error Resume Next
    
    Set varBar = objMenu.Sidebar
    If Not varBar Is Nothing Then
        If (varBar.ItemId = ItemId) Or (varBar.BreakId = ItemId) Then
        
            Set SearchRecursive = varBar
            Exit Function
        
        End If
        
        Set varBar = Nothing
    End If
        
    Set varObj = objMenu.Child
    
    If TypeOf objMenu Is SystemMenu Then Set varSys = objMenu
    
    Do While Not varObj Is Nothing
        
        If (varObj.ItemId = ItemId) Then
            If (hWnd <> 0&) And (Not varSys Is Nothing) Then
                If (hWnd = varSys.hWnd) Then
                    
                    Set SearchRecursive = varObj
                    Exit Do
                
                End If
            Else
                Set SearchRecursive = varObj
                Exit Do
            End If
        End If
        
        If (varObj.Submenu.Count > 0&) Then
            Set varFind = SearchRecursive(varObj.Submenu, ItemId, hWnd)
            
            If Not varFind Is Nothing Then
                Set SearchRecursive = varFind
                Exit Do
            End If
            
        End If
        
        Set varObj = varObj.NextItem
    
    Loop

    Set varFind = Nothing
    Set varObj = Nothing
    Set varSys = Nothing

End Function

Public Function ItemByCaption(ByVal Caption As String, Optional ByVal Menu As Object) As MenuItem
    Dim itemEnum As MenuItem, _
        lpObject As Object
        
        
    If Menu Is Nothing Then
        For Each lpObject In Me
            For Each itemEnum In lpObject
                If (itemEnum.Caption = Caption) Or _
                    (itemEnum.CaptionPlain = Caption) Then
                    
                    Set ItemByCaption = itemEnum
                    Exit Function
                    
                End If
            Next itemEnum
        Next lpObject
        
        Set itemEnum = Nothing
        Set ItemByCaption = Nothing
        
        Exit Function
    End If

    For Each itemEnum In Menu
        If (itemEnum.Caption = Caption) Or _
            (itemEnum.CaptionPlain = Caption) Then
            
            Set ItemByCaption = itemEnum
            Exit For
            
        End If
    Next itemEnum

    Set itemEnum = Nothing

End Function

Public Function ItemByKey(varIndexKey, Optional ByVal Menu As Object) As MenuItem
    Dim lpBar As Menubar, _
        lpMenu As Submenu, _
        lpSys As SystemMenu
        
    Dim objEnum As Object, _
        lpFound As MenuItem
        
    If Menu Is Nothing Then
        For Each objEnum In m_Col
            
            If TypeOf objEnum Is Menubar Then
                Set lpBar = objEnum
                Set lpFound = lpBar(varIndexKey)
                Set lpBar = Nothing
                
            ElseIf TypeOf objEnum Is Submenu Then
                Set lpMenu = objEnum
                Set lpFound = lpMenu(varIndexKey)
                Set lpMenu = Nothing
            
            ElseIf TypeOf objEnum Is SystemMenu Then
                Set lpSys = objEnum
                Set lpFound = lpSys(varIndexKey)
                Set lpSys = Nothing
            
            End If
                        
            If Not lpFound Is Nothing Then
                Set ItemByKey = lpFound
                Set lpFound = Nothing
                
                Exit Function
            End If
        
        Next objEnum
    Else
        
        If TypeOf Menu Is Menubar Then
            Set lpBar = Menu
            Set lpFound = lpBar(varIndexKey)
            Set lpBar = Nothing
            
        ElseIf TypeOf Menu Is Submenu Then
            Set lpMenu = Menu
            Set lpFound = lpMenu(varIndexKey)
            Set lpMenu = Nothing
        
        ElseIf TypeOf Menu Is SystemMenu Then
            Set lpSys = Menu
            Set lpFound = lpSys(varIndexKey)
            Set lpSys = Nothing
        
        End If
                    
        If Not lpFound Is Nothing Then
            Set ItemByKey = lpFound
            Set lpFound = Nothing
        End If
        
    End If

End Function

Public Function GetMenubar(ByVal hWnd As Long) As Menubar
    Dim varObj As Object, _
        varMenu As Menubar
            
    If (m_Active = False) Then Exit Function
    
    For Each varObj In m_Col
        If TypeOf varObj Is Menubar Then
            Set varMenu = varObj
            
            If varMenu.hWnd = hWnd Then
                Exit For
            End If
            
            Set varMenu = Nothing
        
        End If
                
    Next varObj
    
    Set GetMenubar = varMenu
    Set varMenu = Nothing
    Set varObj = Nothing
    
End Function

'' Designer objects are not persisted

Public Sub SaveProperties(PropBag As PropertyBag, Optional ByVal Prefix As String)
    On Error Resume Next
    
    Dim varObj As Object, _
        i As Long
        
    Dim sPrefix As String
    
    If (Prefix <> "") Then sPrefix = Prefix + "."
        
    i = MenuCount
    
    PropBag.WriteProperty sPrefix + "MenuDrawStyle", m_MenuDrawStyle
    PropBag.WriteProperty sPrefix + "GlobalFont", g_SysMenuFont
    PropBag.WriteProperty sPrefix + "Count", i
    
    i = 1&
    For Each varObj In Me
        
        If (TypeOf varObj Is Submenu) Or _
            (TypeOf varObj Is Menubar) Or _
            (TypeOf varObj Is SystemMenu) Then
    
            If (TypeOf varObj Is Submenu) Then
                PropBag.WriteProperty sPrefix + "ItemType" & i, (1&)
                
            ElseIf (TypeOf varObj Is SystemMenu) Then
                PropBag.WriteProperty sPrefix + "ItemType" & i, (2&)
            
            Else
                PropBag.WriteProperty sPrefix + "ItemType" & i, (3&)
            
            End If
            
            varObj.SaveProperties PropBag, sPrefix + "Menu" & i
            i = i + 1&
        End If
    
    Next varObj

End Sub

Public Sub LoadProperties(PropBag As PropertyBag, Optional Prefix As String)
    On Error Resume Next
    
    Dim varObj As Object, _
        i As Long, _
        j As Long
        
    Dim t As Long, _
        sKey As String
        
    Dim sPrefix As String
    
    If (Prefix <> "") Then sPrefix = Prefix + "."
        
    m_MenuDrawStyle = PropBag.ReadProperty(sPrefix + "MenuDrawStyle", m_MenuDrawStyle)
    Set g_SysMenuFont = PropBag.ReadProperty(sPrefix + "GlobalFont", g_SysMenuFont)
    i = PropBag.ReadProperty(sPrefix + "Count", i)
    
    For j = 1& To i
        t = 0&
        t = PropBag.ReadProperty(sPrefix + "ItemType" & j, t)
    
        Select Case t
        
            Case 1:
                Set varObj = New Submenu
            
            Case 2:
                Set varObj = New SystemMenu
            
            Case 3:
                Set varObj = New Menubar
                
        End Select
        
        If Not varObj Is Nothing Then
            sKey = ""
            
            varObj.LoadProperties PropBag, sPrefix + "Menu" & j
            sKey = varObj.Key
            
            If (sKey <> "") Then
                m_Col.Add varObj, sKey
            Else
                m_Col.Add varObj
            End If
            
            Set varObj = Nothing
        End If
        
    Next j
    
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    SaveProperties PropBag
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    LoadProperties PropBag
End Sub

Private Sub Class_Initialize()
    Set m_Col = New Collection
    Set m_PrefixLib = New Accelerators
    
    m_PrefixLib.IsPrefixLib = True


End Sub

Private Sub Class_Terminate()
    
    If (g_MenuCol Is Me) Then
        Set ScaleTool = Nothing
        Set g_WindowList = Nothing
        Set g_Handles = Nothing
        Set g_SysMenuFont = Nothing
    End If
    
    Set m_PrefixLib = Nothing
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

