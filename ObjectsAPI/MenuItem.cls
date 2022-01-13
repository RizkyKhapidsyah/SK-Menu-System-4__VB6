VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"SubMenu"
''' GUInerd Standard Menu System
''' Version 4.1

''' Objects/API Dll

''' Menu Item Class Object


''' ****************************************************************************
''' *** Check the end of the file for copyright and distribution information ***
''' ****************************************************************************

Option Explicit

'' Line Information
'' For multiline functionality.

Private HasLines As Boolean
Private Lines() As String
Private SizeLines As SIZEAPI
Private AccelLine As Long


Public Enum SeparatorTypeConstants
    mstNormal = &H0&
    mstCaption = &H1&
    mstBreak = &H2&
    mstBarBreak = &H4&
End Enum

Public Enum SelectStateConstants
    selNot = 0&
    selSelected = 1&
    selHotLight = 2&
End Enum

'''' Object Properties

'' Accelerator command
Private WithEvents m_Accelerator As Accelerator
Attribute m_Accelerator.VB_VarHelpID = -1

'' Parent Menu
Private m_Parent As Object

'' Previous Item
Private m_PrevItem As MenuItem

'' Next Item
Private m_NextItem As MenuItem

'' Item's Submenu
Private m_SubMenu As Submenu

'' Item Icon (Unchecked, default)
Private m_Picture As StdPicture

'' Item Checked Icon
Private m_CheckedPicture As StdPicture

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

'' ItemInfo Structure

Private m_ItemInfo As MENUITEMINFO

'' Item Id (assigned on Initialize or LoadProperties)
Private m_ItemId As Long

'' Current physical position of menu item
Private m_Index As Long

'' Item collection key
Private m_Key As String

''' Visual Properties

'' Item Caption
Private m_Caption As String

'' Item Checkmarks
Private m_Checked As Boolean
Private m_RadioGroup As Boolean
Private m_Checkmark As Boolean

'' Visibility/Functionality

Private m_Visible As Boolean
Private m_Enabled As Boolean

Private m_Grayed As Boolean

'' Default execution
Private m_Default As Boolean

'' Separator-type item
Private m_Separator As Boolean
Private m_SeparatorType As SeparatorTypeConstants

'' Miscellaneous Data
Private m_Tag As String

Private m_ItemData As Long

''' Position

Private m_RightAlign As Boolean

''' Language

Private m_RightToLeft As Boolean

'' Click event
Public Event Click()

'' RightClick (raised when this item is right-clicked)
Public Event RightClick()


Public Property Get PrevItem() As MenuItem
    Set PrevItem = m_PrevItem
End Property

Friend Property Set PrevItem(ByVal vData As MenuItem)
    Set m_PrevItem = vData
End Property

Public Property Get NextItem() As MenuItem
    Set NextItem = m_NextItem
End Property

Friend Property Set NextItem(ByVal vData As MenuItem)
    Set m_NextItem = vData
End Property

Public Property Get Parent() As Object
    Set Parent = m_Parent
End Property

Friend Property Set Parent(ByVal vData As Object)
    On Error Resume Next
    
    Set m_Parent = Nothing
    
    If (TypeOf vData Is Menubar) Or (TypeOf vData Is Submenu) Or _
        (TypeOf vData Is SystemMenu) Then
        
        Set m_Parent = vData
        
        '' This is important!
        Set m_SubMenu.Parent = vData
        
        Set m_Accelerator.Table = m_Parent.Accelerators
        
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

Public Property Get Submenu() As Submenu
    Set Submenu = m_SubMenu
End Property

Friend Property Set Submenu(vData As Submenu)
    On Error Resume Next
    
    If Not vData Is Nothing Then
        Set m_SubMenu = vData
        SetCtrlState True
    End If
    
End Property

Public Property Get TopMostParent() As Object
    If Not m_Parent Is Nothing Then
        Set TopMostParent = m_Parent.TopMostParent
    End If
End Property

Public Property Get Accelerator() As Accelerator
    On Error Resume Next
    
    Set Accelerator = m_Accelerator
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

Public Property Get CheckedPicture() As StdPicture
    Set CheckedPicture = m_CheckedPicture
End Property

Public Property Set CheckedPicture(vData As StdPicture)
    Set m_CheckedPicture = vData
    SetCtrlState
End Property

Public Property Get ItemId() As Long
    ItemId = m_ItemId
    
End Property

Friend Property Let ItemId(ByVal vData As Long)
    On Error Resume Next
    
    m_ItemId = vData
    m_SubMenu.ItemId = vData
    m_ItemInfo.wID = vData
    
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

Public Property Get Tag() As String
    Tag = m_Tag
End Property

Public Property Let Tag(ByVal vData As String)
    If (m_Tag <> vData) Then
        m_Tag = vData
    End If
End Property

Public Property Get ItemData() As Long
    ItemData = m_ItemData
End Property

Public Property Let ItemData(ByVal vData As Long)
    If (m_ItemData <> vData) Then
        m_ItemData = vData
    End If
End Property

Public Property Get ItemData_String() As String
    ItemData_String = ExtractString(m_ItemData, Using_Unicode)
    
End Property

Public Property Get CaptionPlain() As String
    CaptionPlain = GetPlainCaption(m_Caption)
End Property

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = 0
Attribute Caption.VB_MemberFlags = "200"
    Caption = m_Caption
End Property

Public Property Let Caption(vData As String)
    Dim x As Integer, v As String
        
    If (m_Caption <> vData) Then
        m_Caption = vData
        
        If (m_Caption = "-") Then
            m_Separator = True
            m_SeparatorType = mstNormal
        End If
        
        HasLines = False
        SetCtrlState
    End If
    
End Property

Public Property Get Checkmark() As Boolean
    Checkmark = m_Checkmark
End Property

Public Property Let Checkmark(ByVal vData As Boolean)
    
    If m_Checkmark <> vData Then
        
        If m_CheckedPicture Is Nothing Then
            If (m_RadioGroup = True) Then
                Set m_CheckedPicture = LoadResPicture(101, vbResBitmap)
            Else
                Set m_CheckedPicture = LoadResPicture(102, vbResBitmap)
            End If
            
        End If
        
        m_Checkmark = vData
        SetCtrlState
    End If
    
End Property

Public Property Get RadioGroup() As Boolean
    RadioGroup = m_RadioGroup
End Property

Public Property Let RadioGroup(ByVal vData As Boolean)
    
    If (m_RadioGroup <> vData) Then
        
        If (vData = True) Then
            Set m_CheckedPicture = LoadResPicture(101, vbResBitmap)
        Else
            Set m_CheckedPicture = LoadResPicture(102, vbResBitmap)
        End If
            
        m_RadioGroup = vData
        SetCtrlState
    End If
    
End Property

Public Property Let Checked(ByVal vData As Boolean)
    Dim objSubmenu As Submenu, _
        objMenubar As Menubar
    
    If (m_Checked <> vData) Then
        If (m_RadioGroup = True) And (vData = True) Then
            If Not m_Parent Is Nothing Then
                If TypeOf m_Parent Is Submenu Then
                    Set objSubmenu = m_Parent
                    objSubmenu.ClearRadioGroup
                ElseIf TypeOf m_Parent Is Menubar Then
                    Set objMenubar = m_Parent
                    objMenubar.ClearRadioGroup
                End If
            End If
        End If
    
        m_Checked = vData
        SetCtrlState True
    End If
End Property

Public Property Get Checked() As Boolean
    Checked = m_Checked
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal vData As Boolean)
    If (m_Enabled <> vData) Then
        m_Enabled = vData
        SetCtrlState True
    End If
End Property

Public Property Get Grayed() As Boolean
    Grayed = m_Grayed
End Property

Public Property Let Grayed(ByVal vData As Boolean)
    If (m_Grayed <> vData) Then
        m_Grayed = vData
        SetCtrlState True
    End If
End Property

Public Property Get Visible() As Boolean
    Visible = m_Visible
End Property

Public Property Let Visible(ByVal vData As Boolean)
    If (m_Visible <> vData) Then
        m_Visible = vData
        
        RefreshItem m_ItemId
        
        If (m_Visible = True) Then
            SetCtrlState
        Else
            If TypeOf m_Parent Is Menubar Then
                If (m_Parent.hWnd <> 0&) Then
                    DrawMenuBar m_Parent.hWnd
                End If
            End If
        End If
    End If
End Property

Public Property Get Default() As Boolean
    Default = m_Default
End Property

Public Property Let Default(ByVal vData As Boolean)
    If (m_Default <> vData) Then
        m_Default = vData
        SetCtrlState
    End If
End Property

Public Property Get Separator() As Boolean
    Separator = m_Separator
End Property

Public Property Let Separator(ByVal vData As Boolean)
    If (m_Separator <> vData) Then
        m_Separator = vData
        SetCtrlState
    End If
End Property

Public Property Get SeparatorType() As SeparatorTypeConstants
    SeparatorType = m_SeparatorType
End Property

Public Property Let SeparatorType(ByVal vData As SeparatorTypeConstants)
    If (m_SeparatorType <> vData) Then
        If (vData <> mstNormal) And (m_Separator = False) Then
            m_Separator = True
        End If
        
        If (vData = mstCaption) Then
            With m_Visual
                .ItemBackground = GetSysColor(COLOR_INFOBK)
                .TextAlign = taCenter
            End With
        End If
        
        m_SeparatorType = vData
        SetCtrlState
    End If
End Property

Public Property Get RightToLeft() As Boolean
    RightToLeft = m_RightToLeft
End Property

Friend Property Let RightToLeft(ByVal vData As Boolean)
    If (m_RightToLeft <> vData) Then
        m_RightToLeft = vData
        SetCtrlState
    End If
End Property

Public Property Get RightAlign() As Boolean
    RightAlign = m_RightAlign
End Property

Public Property Let RightAlign(ByVal vData As Boolean)
    If (m_RightAlign <> vData) Then
        m_RightAlign = vData
        SetCtrlState
    End If
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
    Set Search = g_MenuCol.SearchRecursive(m_SubMenu, varIndexKey)
End Function
   
Public Function CopyItem(Optional ByVal Item As MenuItem, Optional ByVal IsSource As Boolean) As MenuItem
    Dim varCopy As MenuItem, _
        PropBag As New PropertyBag
    
    If Not Item Is Nothing Then
        Set varCopy = Item
    Else
        Set varCopy = New MenuItem
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
    m_Default = PropBag.ReadProperty(sPrefix + "Default", m_Default)
    m_Enabled = PropBag.ReadProperty(sPrefix + "Enabled", m_Enabled)
    m_Grayed = PropBag.ReadProperty(sPrefix + "Grayed", m_Grayed)
    m_Visible = PropBag.ReadProperty(sPrefix + "Visible", m_Visible)
    m_Checkmark = PropBag.ReadProperty(sPrefix + "Checkmark", m_Checkmark)
    m_RadioGroup = PropBag.ReadProperty(sPrefix + "RadioGroup", m_RadioGroup)
    m_Checked = PropBag.ReadProperty(sPrefix + "Checked", m_Checked)
    
    m_RightToLeft = PropBag.ReadProperty(sPrefix + "RightToLeft", m_RightToLeft)
    
    m_RightAlign = PropBag.ReadProperty(sPrefix + "RightAlign", m_RightAlign)
                
    m_Tag = PropBag.ReadProperty(sPrefix + "Tag", m_Tag)
                
    m_Separator = PropBag.ReadProperty(sPrefix + "Separator", m_Separator)
    m_SeparatorType = PropBag.ReadProperty(sPrefix + "SeparatorType", m_SeparatorType)
                                           
    Set m_Accelerator = PropBag.ReadProperty(sPrefix + "Accelerator", m_Accelerator)
    
    Set m_Picture = PropBag.ReadProperty(sPrefix + "Picture", m_Picture)
    Set m_CheckedPicture = PropBag.ReadProperty(sPrefix + "CheckedPicture", m_CheckedPicture)
                    
    m_SubMenu.LoadProperties PropBag, sPrefix + "SubMenu"

    HasLines = False

End Sub
   
Public Sub SaveProperties(PropBag As PropertyBag, Optional Prefix As String)
    Dim sPrefix As String
    
    If (Prefix <> "") Then sPrefix = Prefix + "."

    PropBag.WriteProperty sPrefix + "Visual", m_Visual
    
    PropBag.WriteProperty sPrefix + "Key", m_Key
    PropBag.WriteProperty sPrefix + "Caption", m_Caption
    PropBag.WriteProperty sPrefix + "Default", m_Default
    PropBag.WriteProperty sPrefix + "Enabled", m_Enabled
    PropBag.WriteProperty sPrefix + "Grayed", m_Grayed
    PropBag.WriteProperty sPrefix + "Visible", m_Visible
    PropBag.WriteProperty sPrefix + "Checkmark", m_Checkmark
    PropBag.WriteProperty sPrefix + "RadioGroup", m_RadioGroup
    PropBag.WriteProperty sPrefix + "Checked", m_Checked
    
    PropBag.WriteProperty sPrefix + "RightToLeft", m_RightToLeft
    PropBag.WriteProperty sPrefix + "RightAlign", m_RightAlign
                
    PropBag.WriteProperty sPrefix + "Tag", m_Tag
                
    PropBag.WriteProperty sPrefix + "Separator", m_Separator
    PropBag.WriteProperty sPrefix + "SeparatorType", m_SeparatorType
                    
    PropBag.WriteProperty sPrefix + "Accelerator", m_Accelerator
    
    PropBag.WriteProperty sPrefix + "Picture", m_Picture
    PropBag.WriteProperty sPrefix + "CheckedPicture", m_CheckedPicture
    
    m_SubMenu.SaveProperties PropBag, sPrefix + "SubMenu"
    
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
    
    If (m_Grayed = True) Then
        fState = (fState Or MFS_GRAYED)
    End If
    
    If (m_Checked = True) Then
        If (m_RadioGroup = False) Then
            m_Checkmark = True
        Else
            m_Checkmark = False
        End If
    End If
    
    If (m_RadioGroup = True) Or (m_Checkmark = True) Then
        
        If (m_RadioGroup = True) Then
            fType = MFT_RADIOGROUP
        End If
        
        If (m_Checked = True) Then
            fState = fState + MFS_CHECKED
        Else
            fState = fState + MFS_UNCHECKED
        End If
        
    End If
    
    If (m_Caption = "-") Then
        m_Separator = True
        m_SeparatorType = mstNormal
    End If
            
    If (m_Separator = True) Then
        Select Case m_SeparatorType
        
            Case mstNormal:
                fType = MFT_SEPARATOR + MFT_OWNERDRAW
            
            Case mstBreak:
                fType = MFT_MENUBREAK + MFT_OWNERDRAW
            
            Case mstBarBreak:
                fType = MFT_MENUBARBREAK + MFT_OWNERDRAW
            
            Case mstCaption:
                fType = MFT_OWNERDRAW
                
                If (m_SubMenu.AnyItemVisible = False) Then
                    fState = fState Or MFS_DISABLED
                End If
                
        End Select
    Else
        fType = fType + MFT_OWNERDRAW
    End If
        
    If (m_Visual.TextAlign = taRight) Or (m_RightAlign = True) Then
        fType = fType + MFT_RIGHTJUSTIFY
    End If
        
    If (m_Default = True) Then
        fState = fState + MFS_DEFAULT
    End If
    
    If (m_Grayed = True) Then
        fType = (fType Or MFT_SEPARATOR)
    End If
    
    m_ItemInfo.fMask = MIIM_STATE + MIIM_TYPE
    
    m_ItemInfo.fType = fType
    m_ItemInfo.fState = fState
    
    m_ItemInfo.dwTypeData = 0&
    
    hParent = m_Parent.hMenu
    
    If (m_SubMenu.hMenu <> 0&) Then
       m_ItemInfo.fMask = m_ItemInfo.fMask + MIIM_SUBMENU
       m_ItemInfo.hSubMenu = m_SubMenu.hMenu
    End If
    
    SetMenuItemInfo_API hParent, m_ItemId, False, m_ItemInfo
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
        
    If (m_ItemInfo.fState And MFS_CHECKED) Then
        m_Checked = True
        m_Checkmark = True
    Else
        m_Checked = False
    End If

    If (m_ItemInfo.fType And MFT_RADIOGROUP) Then
        m_RadioGroup = True
        m_Checkmark = False
    End If

    If (m_Visual.TextAlign <> taCenter) Then
        If (m_ItemInfo.fType And MFT_RIGHTJUSTIFY) Then
            m_Visual.TextAlign = taRight
        Else
            m_Visual.TextAlign = taLeft
        End If
    End If
    
    
    If (m_ItemInfo.fState And MFS_DEFAULT) Then
        m_Default = True
    Else
        m_Default = False
    End If
    
    '' Separator ?
    
    If (m_ItemInfo.fType And (MFT_MENUBREAK + MFT_MENUBARBREAK)) Then
        m_Separator = True
        
        If (m_ItemInfo.fType And MFT_MENUBREAK) Then
            m_SeparatorType = mstBarBreak
        Else
            m_SeparatorType = mstBreak
        End If
    
    ElseIf (m_ItemInfo.fType And MFT_SEPARATOR) Then
        
        m_SeparatorType = mstNormal
        m_Separator = True
        
    Else
        
        If (m_SeparatorType And mstCaption) Then
            m_SeparatorType = mstCaption
            m_Separator = True
        
        Else
            m_Separator = False
            m_SeparatorType = 0&
        End If
                
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

Friend Sub Destroy(Optional ByVal NoDeleteMenu As Boolean)
    
    If Not m_Parent Is Nothing Then
        
        If NoDeleteMenu = False Then
            DeleteMenu m_Parent.hMenu, m_ItemId, MF_BYCOMMAND
        End If
    End If
    
End Sub

Private Sub Recreate()
    Dim hTop As Long, _
        i As Long, _
        j As Long
    
    On Error Resume Next
    
    If (TypeOf m_Parent Is Menubar) Then
        RecreateMenu m_Parent
    ElseIf (TypeOf m_Parent Is Submenu) Or (TypeOf m_Parent Is SystemMenu) Then
        RefreshItem m_ItemId
        SetMenuItemInfo_API m_Parent.hMenu, m_ItemId, False, m_ItemInfo
        
        i = MenuMaxImageWidth(m_Parent, j)
        SetImageMax m_Parent, i, j
        
    End If
    
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
    
    m_Enabled = True
    
    m_Visible = True
    m_ItemId = GetNewCtrlId
            
    Set m_Visual = New ItemVisualProperties
    Set m_Visual.Item = Me
    
    Set m_Accelerator = New Accelerator
    m_Accelerator.ItemId = m_ItemId
            
    Set m_SubMenu = New Submenu
    
    Set m_SubMenu.Parent = m_Parent
    
    m_SubMenu.ItemId = m_ItemId
    
End Sub

Private Sub Class_Terminate()
    
    Erase Lines
    
    Destroy
    
    Set m_SubMenu = Nothing
    Set m_Accelerator = Nothing
    
    Set m_Picture = Nothing
    Set m_CheckedPicture = Nothing
    
    Set m_Visual = Nothing
    
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





