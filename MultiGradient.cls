VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MultiGradient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Col As Collection

'' specify fCopyTo = True to copy data into vData,
'' or False (default) to copy data from vData into this instance

Public Function Copy(vData As MultiGradient, Optional ByVal fCopyTo As Boolean)
    Dim varColor As Variant
        
    If vData Is Nothing Then
        If (fCopyTo = True) Then
            Set vData = New MultiGradient
        End If
    End If
    
    If (fCopyTo = True) Then
        If (vData.Count <> 0&) Then vData.Clear
        
        For Each varColor In m_Col
            vData.Add varColor
        Next varColor
    Else
        Me.Clear
        
        For Each varColor In vData
            m_Col.Add varColor
        Next varColor
    End If
    
End Function

Friend Function GetColorArray() As Long()
    Dim b() As Long, _
        i As Long, _
        v
        
    For Each v In m_Col
        ReDim Preserve b(0 To i)
        b(i) = GetActualColor(v)
        i = i + 1
    Next v

    GetColorArray = b
    
End Function

Public Sub Add(ByVal crColor As OLE_COLOR, Optional Key, Optional InsertBefore)
    Dim var
    
    var = crColor
    
    If IsMissing(Key) = True Then
        If IsMissing(InsertBefore) = True Then
            m_Col.Add var
        Else
            m_Col.Add var, , InsertBefore
        End If
    Else
        If IsMissing(InsertBefore) = True Then
            m_Col.Add var, Key
        Else
            m_Col.Add var, Key, InsertBefore
        End If
    End If
        
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_Col.[_NewEnum]
End Function

Public Sub Remove(varIndexKey)
    m_Col.Remove varIndexKey
End Sub

Public Sub Clear()
    Set m_Col = New Collection
End Sub

Public Property Get Count() As Long
    Count = m_Col.Count
End Property

Public Property Get Color(varIndexKey) As OLE_COLOR
Attribute Color.VB_UserMemId = 0
Attribute Color.VB_MemberFlags = "200"
    Dim var
    
    var = m_Col(varIndexKey)
    Color = var
    
End Property

Public Sub SaveProperties(PropBag As PropertyBag, Optional Prefix As String)
    Dim v As Variant, _
        val As Long, _
        i As Long, _
        sPrefix As String
        
    If (Prefix <> "") Then
        sPrefix = Prefix + "."
    End If
        
    i = m_Col.Count
    PropBag.WriteProperty sPrefix + "Count", i
        
    i = 1&
    For Each v In m_Col
        val = v
        
        PropBag.WriteProperty sPrefix + "Value" & i, val
    Next v
    
End Sub

Public Sub LoadProperties(PropBag As PropertyBag, Optional Prefix As String)
    Dim v As Variant, _
        val As Long, _
        i As Long, _
        j As Long, _
        sPrefix As String
        
    Set m_Col = Nothing
    Set m_Col = New Collection
        
    If (Prefix <> "") Then
        sPrefix = Prefix + "."
    End If
    
    i = 0&
    i = PropBag.ReadProperty(sPrefix + "Count", i)

    For j = 1 To i
    
        val = PropBag.ReadProperty(sPrefix + "Value" & j, val)
        v = val
        
        m_Col.Add v
    Next j

End Sub

Private Sub Class_Initialize()
    Set m_Col = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_Col = Nothing
End Sub

