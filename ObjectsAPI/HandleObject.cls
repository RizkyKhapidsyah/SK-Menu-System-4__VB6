VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HandleObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' GUInerd Standard Menu System
''' Version 4.1

''' Objects/API Dll

''' Window Handle Class


''' ****************************************************************************
''' *** Check the end of the file for copyright and distribution information ***
''' ****************************************************************************

Option Explicit

Public Key As String

Private m_hWnd As Long
Private m_hMenu As Long

Private m_MenuId As Long

Private m_NewProc As Long
Private m_OldProc As Long

Private m_PaintDC As Long

Private m_hWndStatic As Long

Private m_Interface As Object

Private m_References As Long

Public Property Get References() As Long
    References = m_References
End Property

Public Sub AddReference()
    m_References = m_References + 1
End Sub

Public Sub RemoveReference()
    m_References = m_References - 1
    If (m_References = 0&) Then Reset
End Sub

Public Sub CreateDC()
    Dim dDC As Long
    
    If (m_hWnd = 0&) Then Exit Sub
    
    If (m_PaintDC <> 0&) Then
        DeleteDC m_PaintDC
        m_PaintDC = 0&
    End If
    
    dDC = GetWindowDC(m_hWnd)
    m_PaintDC = CreateCompatibleDC(dDC)
    ReleaseDC m_hWnd, dDC
        
End Sub

Public Property Get Interface() As Object
    Set Interface = m_Interface
End Property

Public Property Set Interface(ByVal vData As Object)
    Set m_Interface = vData
End Property

Public Property Get hWnd() As Long
    hWnd = m_hWnd
End Property

Public Property Get hWndStatic() As Long
    hWndStatic = m_hWndStatic
End Property

Public Property Get hMenu() As Long
    hMenu = m_hMenu
End Property

Public Property Get OldProc() As Long
    OldProc = m_OldProc
End Property

Public Property Get NewProc() As Long
    NewProc = m_NewProc
End Property

Public Property Get PaintDC() As Long
    PaintDC = m_PaintDC
End Property

Public Sub SetHandle(ByVal hWndNew As Long, Optional ByVal hProcNew As Long, Optional ByVal hMenuNew As Long, Optional ByVal hWndChild As Long)

    If (m_hWnd <> 0&) Then Reset
    
    m_hWnd = hWndNew
    
    If (hProcNew <> 0&) Then
        m_OldProc = SetWindowLong(m_hWnd, GWL_WNDPROC, hProcNew)
        m_NewProc = hProcNew
    End If
    
    m_References = 1&
    
    m_hMenu = hMenuNew
    m_hWndStatic = hWndChild
    
    Key = "_H" + Hex(m_hWnd)
    CreateDC
        
End Sub

Public Sub Reset()

    If (m_hWnd = 0&) Then Exit Sub
    
    If (m_PaintDC <> 0&) Then
        DeleteDC m_PaintDC
        m_PaintDC = 0&
    End If
    
    If (m_OldProc <> 0&) Then
        SetWindowLong m_hWnd, GWL_WNDPROC, m_OldProc
    
        m_OldProc = 0&
        m_NewProc = 0&
    End If
    
    m_hWnd = 0&
    m_hWndStatic = 0&
    
    m_hMenu = 0&
    
    m_References = 0&
    
    Key = ""
    
End Sub

Private Sub Class_Terminate()
    If (m_hWnd <> 0&) Then Reset
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






