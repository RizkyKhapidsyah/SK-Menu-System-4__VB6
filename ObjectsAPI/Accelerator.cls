VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Accelerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
''' GUInerd Standard Menu System
''' Version 4.1

''' Objects/API Dll

''' Menu Accelerator Class Object


''' ****************************************************************************
''' *** Check the end of the file for copyright and distribution information ***
''' ****************************************************************************

Option Explicit

Public Enum AcceleratorShiftStates
    acNone = &H0&
    
    acShift = &H1&
    acCtrl = &H2&
    acAlt = &H4&
    
    '' Extended Shift Keys
    acLShift = &H10&
    acLCtrl = &H20&
    acLAlt = &H40&
    
    acRShift = &H100&
    acRCtrl = &H200&
    acRAlt = &H400&
    
    acIsExtended = &H770&
End Enum

Private m_SysMenu As Boolean

Private m_ExtendedKeys As Boolean

Private m_KeyCode As Long

Private m_Shift As AcceleratorShiftStates

Private m_CommandKey As Long

Private m_ItemId As Long

Private m_Table As Accelerators

Private m_Word As String

Public Event AcceleratorChange()

Public Property Get ExtendedKeys() As Boolean
    ExtendedKeys = CBool(g_MenuCol.AcceleratorStyle And sbExtendedKeys)
End Property

Public Property Get KeyCode() As Long
    KeyCode = m_KeyCode
End Property

Public Property Let KeyCode(ByVal vData As Long)
    m_KeyCode = vData
End Property

Public Property Get CommandKey() As Long
    CommandKey = m_CommandKey
End Property

Public Property Let CommandKey(ByVal vData As Long)
    m_CommandKey = vData
    ReValidate
End Property

Public Property Get Shift() As AcceleratorShiftStates
    Shift = m_Shift
End Property

Public Property Let Shift(ByVal vData As AcceleratorShiftStates)
    m_Shift = vData
    ReValidate
End Property

Public Property Get ItemId() As Long
    ItemId = m_ItemId
End Property

Friend Property Let ItemId(ByVal vData As Long)
    m_ItemId = vData
End Property

Public Property Get IsCommandPrefix() As Boolean
    If (m_ItemId = -1&) Then IsCommandPrefix = True Else _
        IsCommandPrefix = False
End Property

Public Property Get Table() As Accelerators
    Set Table = m_Table
End Property

Friend Property Set Table(vData As Accelerators)
    Set m_Table = vData
    If m_Table Is Nothing Then
        m_SysMenu = False
        Exit Property
    End If
    
    If Not m_Table.Menu Is Nothing Then
        If TypeOf m_Table.Menu Is SystemMenu Then
            m_SysMenu = True
        Else
            m_SysMenu = False
        End If
    End If
End Property

Public Property Get AccelWord() As String
Attribute AccelWord.VB_MemberFlags = "200"
    If m_KeyCode = 0& Then Exit Property
    
    AccelWord = AccelCodeToString(m_KeyCode, m_Shift, m_CommandKey)
End Property

Public Property Let AccelWord(ByVal vData As String)
    Dim varStr As String
    
    If vData <> "" Then
        
        StringToAccelCode vData, m_KeyCode, m_Shift, m_CommandKey
        varStr = AccelCodeToString(m_KeyCode, m_Shift, m_CommandKey)
        
        ReValidate
        
        If (varStr <> m_Word) Then
            m_Word = varStr
            RaiseEvent AcceleratorChange
        End If
    Else
        m_KeyCode = 0&
        m_Shift = 0&
    End If
    
End Property

Public Sub SaveProperties(PropBag As PropertyBag, Optional ByVal Prefix As String)
    Dim sPrefix As String
    
    If (Prefix <> "") Then sPrefix = Prefix + "."
    
    PropBag.WriteProperty sPrefix + "Shift", m_Shift
    PropBag.WriteProperty sPrefix + "KeyCode", m_KeyCode
    PropBag.WriteProperty sPrefix + "CommandKey", m_CommandKey
    PropBag.WriteProperty sPrefix + "ExtendedKeys", m_ExtendedKeys
    PropBag.WriteProperty sPrefix + "ItemId", m_ItemId
        
End Sub

Public Sub LoadProperties(PropBag As PropertyBag, Optional ByVal Prefix As String)
    Dim sPrefix As String
    
    If (Prefix <> "") Then sPrefix = Prefix + "."

    m_Shift = PropBag.ReadProperty(sPrefix + "Shift", m_Shift)
    m_KeyCode = PropBag.ReadProperty(sPrefix + "KeyCode", m_KeyCode)
    m_CommandKey = PropBag.ReadProperty(sPrefix + "CommandKey", m_CommandKey)
    m_ExtendedKeys = PropBag.ReadProperty(sPrefix + "ExtendedKeys", m_ExtendedKeys)
    m_ItemId = PropBag.ReadProperty(sPrefix + "ItemId", m_ItemId)
    
End Sub

Private Sub ReValidate()
    Dim iValidate As AcceleratorValidationResults

    iValidate = ValidateAccelerator(m_KeyCode, m_Shift, m_CommandKey, m_SysMenu, m_Table.IsPrefixLib)
        
    If (iValidate And acError) Then
        PostAcceleratorErrorMsg iValidate
    End If
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    SaveProperties PropBag
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    LoadProperties PropBag
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






