VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Accelerators"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Accel"
Attribute VB_Ext_KEY = "Member0" ,"Accel"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
''' GUInerd Standard Menu System
''' Version 4.1

''' Objects/API Dll

''' Accelerator Table and Collection Class Object


''' ****************************************************************************
''' *** Check the end of the file for copyright and distribution information ***
''' ****************************************************************************

Option Explicit

'local variable to hold collection

Private m_Col As Collection

Private m_Menu As Object

Private m_IsPrefixLib As Boolean

Public Property Get Menu() As Object
    Set Menu = m_Menu
End Property

Friend Property Set Menu(ByVal vData As Object)
    Set m_Menu = vData
End Property

Public Property Get IsPrefixLib() As Boolean
    IsPrefixLib = m_IsPrefixLib
End Property

Friend Property Let IsPrefixLib(ByVal vData As Boolean)
    m_IsPrefixLib = vData
End Property

Public Property Get ExtendedKeys() As Boolean
    ExtendedKeys = CBool(g_MenuCol.AcceleratorStyle And sbExtendedKeys)
End Property

Public Function Add(ByVal AccelWord As String, _
                    Optional Item) As Accelerator
            
    On Error Resume Next
                    
    Dim varObj As MenuItem, _
        varNew As Accelerator
    
    If (IsMissing(Item) = False) And (m_IsPrefixLib = False) Then
        If IsObject(Item) And (TypeOf Item Is MenuItem) Then
            Set varObj = g_MenuCol.FindMenuItem(Item.ItemId)
        ElseIf Not IsObject(Item) Then
            Set varObj = g_MenuCol.FindMenuItem(Item)
        End If
    End If
                
    If Not varObj Is Nothing Then
        Set varNew = varObj.Accelerator
        
        If varNew Is Nothing Then
            Set varNew = New Accelerator
            Set varObj.Accelerator = varNew
        End If
    Else
        Set varNew = New Accelerator
    End If
        
    Set varNew.Table = Me
        
    With varNew
        .AccelWord = AccelWord
        
        If (m_IsPrefixLib = True) Then
            .ItemId = -1&
        End If
    
        If (.ItemId <> -1&) And (.ItemId <> 0&) Then
            m_Col.Add varNew, "_H" + Hex(.ItemId)
        Else
            m_Col.Add varNew
        End If
        
    End With
    
End Function
                    
Public Property Get Item(ByVal Index As Long) As Accelerator
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
    Set Item = m_Col(Index)
End Property

Public Property Get ItemByID(ByVal ItemId As Long) As Accelerator
    If (ItemId <> 0&) And (ItemId <> -1&) Then
        Set ItemByID = m_Col("_H" + Hex(ItemId))
    End If
End Property

Public Property Get Count() As Long
    Count = m_Col.Count
End Property

Public Sub Remove(ByVal Index As Long)
    m_Col.Remove Index
End Sub

Public Sub RemoveByID(ByVal ItemId As Long)
    If (ItemId <> 0&) And (ItemId <> -1&) Then
        m_Col.Remove "_H" + Hex(ItemId)
    End If
End Sub

Public Sub Clear()
    Set m_Col = Nothing
    Set m_Col = New Collection
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_Col.[_NewEnum]
End Property

Public Sub SaveProperties(PropBag As PropertyBag, Optional ByVal Prefix As String)
    Dim sPrefix As String, _
        i As Long
        
    Dim varObj As Accelerator
    
    If (Prefix <> "") Then sPrefix = Prefix + "."
    
    i = m_Col.Count
    PropBag.WriteProperty sPrefix + "Count", i
    
    i = 1&
    
    For Each varObj In m_Col
        varObj.SaveProperties PropBag, sPrefix + "Accelerator" & i
        i = i + 1
    Next varObj
    
End Sub

Public Sub LoadProperties(PropBag As PropertyBag, Optional ByVal Prefix As String)
    Dim sPrefix As String, _
        j As Long, _
        i As Long
    
    Dim varObj As Accelerator
    
    If (Prefix <> "") Then sPrefix = Prefix + "."

    Set m_Col = New Collection
    i = PropBag.ReadProperty(sPrefix + "Count", i)
    
    For j = 1 To i
    
        Set varObj = New Accelerator
        varObj.LoadProperties PropBag, sPrefix + "Accelerator" & j
        
        m_Col.Add varObj
        Set varObj = Nothing
    Next j
    
End Sub

Friend Property Get Col() As Collection
    Set Col = m_Col
End Property

Private Sub Class_WriteProperties(PropBag As PropertyBag)
    SaveProperties PropBag
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    LoadProperties PropBag
End Sub

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set m_Col = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
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






