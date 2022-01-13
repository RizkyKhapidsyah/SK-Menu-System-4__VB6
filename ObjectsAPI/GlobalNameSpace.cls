VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuLibGlobal"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get Menus() As Menus
    Set Menus = g_MenuCol
End Property

Public Property Get MenuItems(varIndexKey) As MenuItem
    Set MenuItems = g_MenuCol.ItemByKey(varIndexKey)
End Property

Public Sub SetActiveMenuSet(ByVal MenuSet As Menus, Optional ByVal MDIConnection As Object)
    
    If Not MenuSet Is Nothing Then
        ChangeActiveMenuSet MenuSet, MDIConnection
    End If
    
End Sub

