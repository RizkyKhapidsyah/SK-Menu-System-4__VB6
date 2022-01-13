VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sidebar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''' GUInerd Standard Menu System
''' Version 4.1

''' Objects/API Dll

''' Sidebar Class Object


''' ****************************************************************************
''' *** Check the end of the file for copyright and distribution information ***
''' ****************************************************************************

Option Explicit

'' Line Info for Drawing

Private HasLines As Boolean
Private Lines() As String
Private SizeLines As SIZEAPI
Private AccelLine As Long

Public Enum SidebarEscapementConstants
    esc90 = &H0
    esc270 = &H1
    
    escWestern = &H0
    escEastern = &H1
End Enum

Public Enum SidebarPositionConstants
    posLeft = &H0
    posRight = &H1
    
    'posTop = &H2
    'posBottom = &H3
End Enum


'''' Object Properties

'' Parent Menu
Private m_Parent As Object

'' Item Icon (Unchecked, default)
Private m_Picture As StdPicture

'' menu item visual effects are stored in a separate class module
'' this is simply because there are so many.

Private WithEvents m_Visual As ItemVisualProperties
Attribute m_Visual.VB_VarHelpID = -1

'''' Non-objects

'' Frozen (controls item refresh)
Private m_Frozen As Boolean

'' When a designer is controlling a menu item
'' this flag is set to True
Private m_DesignMode As Boolean

'' This flag has the designer's menu-design mode
Private m_ParentType As MenuTypeConstants

''' Menu identification

'' Item Info Structures

Private m_ItemInfo As MENUITEMINFO
Private m_BreakInfo As MENUITEMINFO

'' Item Id (assigned on Initialize or LoadProperties)
Private m_ItemId As Long

'' MenuBreak Id (assigned on Initialize)
Private m_BreakId As Long

'' Current physical position of menu item
Private m_Index As Long

'' Item collection key
Private m_Key As String

''' Visual Properties

'' Item Caption
Private m_Caption As String

'' Visibility/Functionality

Private m_Visible As Boolean
Private m_Enabled As Boolean

''' Font, Language and Placement

Private m_RightToLeft As Boolean

Private m_Escapement As SidebarEscapementConstants

Private m_Position As SidebarPositionConstants

Public Event Click()

Public Event RightClick()


Public Property Get Parent() As Object
    Set Parent = m_Parent
End Property

Friend Property Set Parent(ByVal vData As Object)
    On Error Resume Next
    
    Set m_Parent = Nothing
    
    If (TypeOf vData Is Menubar) Or (TypeOf vData Is Submenu) Or _
        (TypeOf vData Is SystemMenu) Then
        
        Set m_Parent = vData
        
        If (TypeOf m_Parent Is Menubar) Then
            m_ParentType = mtcMenubar
        ElseIf (TypeOf m_Parent Is Submenu) Then
            m_ParentType = mtcPopup
        ElseIf (TypeOf m_Parent Is SystemMenu) Then
            m_ParentType = mtcSysmenu
        End If
        
        m_DesignMode = False
        m_RightToLeft = m_Parent.RightToLeft
        
    Else
        Set m_Parent = vData
        m_DesignMode = True
    End If
    
    
End Property

Public Property Get TopMostParent() As Object
    If Not m_Parent Is Nothing Then
        Set TopMostParent = m_Parent.TopMostParent
    End If
End Property

Public Property Get Visual() As ItemVisualProperties
    Set Visual = m_Visual
End Property

Public Property Set Visual(ByVal vData As ItemVisualProperties)
    Set m_Visual = vData
    SetCtrlState
End Property

Public Property Get Picture() As StdPicture
    Set Picture = m_Picture
End Property

Public Property Set Picture(vData As StdPicture)
    Set m_Picture = vData
    SetCtrlState
End Property

Public Property Get ItemId() As Long
    ItemId = m_ItemId
    
End Property

Friend Property Let ItemId(ByVal vData As Long)
    m_ItemId = vData
    m_ItemInfo.wID = vData
    
    SetCtrlState
End Property
        
Public Property Get BreakId() As Long
    BreakId = m_BreakId
    
End Property

Friend Property Let BreakId(ByVal vData As Long)
    m_BreakId = vData
    m_BreakInfo.wID = vData
    
    SetCtrlState
End Property
        
Public Property Get Index() As Long
    Index = m_Index
End Property

Friend Property Let Index(ByVal vData As Long)
    m_Index = vData
End Property

Public Property Let Key(ByVal vData As String)
    m_Key = vData
End Property

Public Property Get Key() As String
    Key = m_Key
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(vData As String)
    
    If (m_Caption <> vData) Then
        m_Caption = vData
        SetCtrlState
    End If
    
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal vData As Boolean)
    m_Enabled = vData
    SetCtrlState True
End Property

Public Property Get Visible() As Boolean
    Visible = m_Visible
End Property

Public Property Let Visible(ByVal vData As Boolean)
    If (m_Visible <> vData) Then
        m_Visible = vData
        
        RefreshItem m_ItemId
    End If
    
End Property

Public Property Get Escapement() As SidebarEscapementConstants
    Escapement = m_Escapement
End Property

Public Property Let Escapement(ByVal vData As SidebarEscapementConstants)
    m_Escapement = vData
    SetCtrlState
End Property

Public Property Get Position() As SidebarPositionConstants
    Position = m_Position
End Property

Public Property Let Position(ByVal vData As SidebarPositionConstants)
    m_Position = vData
    SetCtrlState
End Property

Public Property Get RightToLeft() As Boolean
    RightToLeft = m_RightToLeft
End Property

Friend Property Let RightToLeft(ByVal vData As Boolean)
    m_RightToLeft = vData
    SetCtrlState
End Property

Public Property Get DesignMode() As Boolean
    DesignMode = m_DesignMode
End Property

Friend Property Let DesignMode(ByVal vData As Boolean)
    m_DesignMode = vData
End Property

Public Property Get ParentType() As MenuTypeConstants
    ParentType = m_ParentType
End Property

Friend Property Let ParentType(ByVal vData As MenuTypeConstants)
    m_ParentType = vData
End Property

Public Function Search(ByVal varIndexKey As Variant) As Object
    
    '' We forward this to parent
    Set Search = g_MenuCol.SearchRecursive(m_Parent, varIndexKey)
    
End Function
   
Public Function CopyItem(Optional ByVal Item As Sidebar, Optional ByVal IsSource As Boolean) As Sidebar
    Dim varCopy As Sidebar, _
        PropBag As New PropertyBag
    
    If Not Item Is Nothing Then
        Set varCopy = Item
    Else
        Set varCopy = New Sidebar
    End If
    
    If IsSource = True Then
        SaveProperties PropBag
        varCopy.LoadProperties PropBag
    Else
        varCopy.SaveProperties PropBag
        LoadProperties PropBag
    End If
    
    Set PropBag = Nothing
    Set CopyItem = varCopy
    Set varCopy = Nothing
    
End Function
 
Public Sub LoadProperties(PropBag As PropertyBag, Optional Prefix As String)
    Dim sPrefix As String
    
    If (Prefix <> "") Then sPrefix = Prefix + "."
    
    Set m_Visual = PropBag.ReadProperty(sPrefix + "Visual", m_Visual)
    
    m_Key = PropBag.ReadProperty(sPrefix + "Key", m_Key)
    m_Caption = PropBag.ReadProperty(sPrefix + "Caption", m_Caption)
    m_Enabled = PropBag.ReadProperty(sPrefix + "Enabled", m_Enabled)
    m_Visible = PropBag.ReadProperty(sPrefix + "Visible", m_Visible)
    
    m_RightToLeft = PropBag.ReadProperty(sPrefix + "RightToLeft", m_RightToLeft)
    m_Escapement = PropBag.ReadProperty(sPrefix + "Escapement", m_Escapement)
    m_Position = PropBag.ReadProperty(sPrefix + "Position", m_Position)
                   
    Set m_Picture = PropBag.ReadProperty(sPrefix + "Picture", m_Picture)
                    
End Sub
   
Public Sub SaveProperties(PropBag As PropertyBag, Optional Prefix As String)
    Dim sPrefix As String
    
    If (Prefix <> "") Then sPrefix = Prefix + "."

    PropBag.WriteProperty sPrefix + "Visual", m_Visual
    
    PropBag.WriteProperty sPrefix + "Key", m_Key
    PropBag.WriteProperty sPrefix + "Caption", m_Caption
    PropBag.WriteProperty sPrefix + "Enabled", m_Enabled
    PropBag.WriteProperty sPrefix + "Visible", m_Visible
    
    PropBag.WriteProperty sPrefix + "RightToLeft", m_RightToLeft
    PropBag.WriteProperty sPrefix + "Position", m_Position
    PropBag.WriteProperty sPrefix + "Escapement", m_Escapement
                
    PropBag.WriteProperty sPrefix + "Picture", m_Picture
    
End Sub
   
   
''' Centralized functions for setting and getting certain aspects of a menu
''' item's control state such as the Enabled and Default properties.

' refresh the physical control from data stored in the object.

Friend Sub SetCtrlState(Optional ByVal fNoRecreate As Boolean)
    Dim fState As Long, _
        fType As Long
    
    Dim hParent As Long
    
    On Error Resume Next
    
    If (m_Frozen = True) Then Exit Sub
    
    ''' if there is no parent, this menu item does not exist.
    fType = 0&
    fState = 0&
    
    If m_Parent Is Nothing Then Exit Sub
    
    ''' Set the caption along with the state
    
    If (m_Enabled = False) Then
        fState = MFS_DISABLED
    Else
        fState = MFS_ENABLED
    End If
        
    fType = fType + MFT_OWNERDRAW
        
    m_ItemInfo.fMask = MIIM_STATE + MIIM_TYPE
    
    m_ItemInfo.fType = fType
    m_ItemInfo.fState = fState
    
    m_ItemInfo.dwTypeData = 0&
    
    hParent = m_Parent.hMenu
    
    SetMenuItemInfo_API hParent, m_ItemId, False, m_ItemInfo
    
    m_ItemInfo.fType = MFT_OWNERDRAW + MFT_MENUBREAK
    m_ItemInfo.fState = MFS_DISABLED + MFS_GRAYED
    
    m_ItemInfo.fMask = MIIM_STATE + MIIM_TYPE
    
    SetMenuItemInfo_API hParent, m_BreakId, False, m_ItemInfo
    
    If (fNoRecreate = False) Then Recreate
    
End Sub

' Refresh the state of the control object with that of the physical menu's.

Friend Sub GetCtrlState()
    On Error Resume Next
    
    ''' if there is no parent, this menu item does not exist.
        
    If m_Parent Is Nothing Then Exit Sub
    
    m_ItemInfo.fMask = MIIM_STATE + MIIM_TYPE
    
    GetMenuItemInfo_API m_Parent.hMenu, m_ItemId, False, m_ItemInfo
    
    If (m_ItemInfo.fState And (MFS_DISABLED + MFS_GRAYED)) Then
        m_Enabled = False
    Else
        m_Enabled = True
    End If
        
End Sub

Friend Sub Freeze()
    m_Frozen = True
End Sub

Friend Sub Unfreeze()
    m_Frozen = False
End Sub



Friend Function ExecCmd(ByVal wParam As Long, ByVal uMsg As Long) As Long
    
    If (wParam = m_ItemId) Then
        '' should be =True all the time, but still...
        
        Select Case uMsg
        
            Case WM_MENURBUTTONUP
                RaiseEvent RightClick
            
            Case WM_COMMAND, WM_SYSCOMMAND
                RaiseEvent Click
        
        End Select
    End If
        
    SendCommand m_Parent, wParam, uMsg

End Function


Friend Sub Destroy(Optional ByVal NoDeleteMenu As Boolean)
        
    If Not m_Parent Is Nothing Then
        
        If NoDeleteMenu = False Then
            DeleteMenu m_Parent.hMenu, m_ItemId, 0&
        End If
    End If
    
End Sub


Friend Sub GetLineInfo(lpLines() As String, lpdwAccelLine As Long, lpSize As SIZEAPI, Optional ByVal hDC As Long)
    Dim linesGet() As String, _
        x As Long, _
        y As Long
        
    Dim lpGet As SIZEAPI, _
        dAccel As Long
        
    If (HasLines = False) Then
        On Error Resume Next
        
        x = -1&
        x = UBound(Lines)
        If (x <> -1&) Then Erase Lines
        
        If (hDC = 0&) Then
            hDC = GetDC(0&)
        End If
        
        linesGet = Parse_String(hDC, m_Caption, , m_Visual.WrapLimit, dAccel, ByVal VarPtr(lpGet))
        
        x = UBound(linesGet)
        
        ReDim Lines(0 To x)
        
        For y = 0 To x
            Lines(y) = linesGet(y)
        Next y
        
        Erase linesGet
    
        AccelLine = dAccel
        
        SizeLines.cx = lpGet.cx
        SizeLines.cy = lpGet.cy
        
        x = -1&
        x = UBound(Lines)
        
        If (x <> -1&) Then
            HasLines = True
        End If
        
    End If

    x = -1&
    x = UBound(Lines)
    
    If (HasLines = True) Then
        ReDim lpLines(0 To x)
        
        For y = 0 To x
            lpLines(y) = Lines(y)
        Next y
                
        lpdwAccelLine = dAccel
        
        lpSize.cx = SizeLines.cx
        lpSize.cy = SizeLines.cy
    End If
    
End Sub


Private Sub Recreate()
    
    RefreshItem m_ItemId

End Sub

Private Sub m_Visual_PropertyChanged(ByVal PropertyName As String)

    Select Case PropertyName
    
        Case "UserWidth", "UserHeight", "UserSize", _
             "ImageScaleWidth", "ImageScaleHeight", _
             "ScaleImages", "ImageScale", "Font", _
             "WrapLimit"
             
            Recreate
            
    End Select

End Sub


Private Sub m_Accelerator_AcceleratorChange()
    SetCtrlState
End Sub


Private Sub Class_Initialize()
    On Error Resume Next
        
    m_Frozen = True
    
    m_Enabled = False
    m_Visible = False
    
    m_ItemId = GetNewCtrlId
    m_BreakId = GetNewCtrlId
    

    Set m_Visual = New ItemVisualProperties
    Set m_Visual.Item = Me
    
    '' Initial/Default visual properties are a little bit
    '' different for a sidebar.
    
    '' (for one thing, 'fancy' is standard)
    
    m_Visual.ItemBackground = vbButtonFace
    m_Visual.ItemBkGradient = vbInfoBackground
    
    m_Visual.ItemForeground = vbBlack
    
    m_Visual.SelectBackground = vbButtonFace
    m_Visual.SelectBkGradient = vbInfoBackground
    
    m_Visual.SelectionStyle = mssHotTrack
    
    m_Visual.SelectForeground = vbBlue
    
    m_Visual.TextAlign = taCenter
    
End Sub

Private Sub Class_Terminate()
    
    Erase Lines
    Destroy
    
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







