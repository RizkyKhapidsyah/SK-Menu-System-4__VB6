VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ItemVisualProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum TextAlignConstants
    taLeft = TA_LEFT
    taRight = TA_RIGHT
    taCenter = TA_CENTER
End Enum

Public Enum PictureVAlignConstants
    ppTop = 0&
    ppCenter = 1
    ppBottom = 2
End Enum

'' These may all be combined.

Public Enum SelectionStyleConstants
    mssDefault = &H0&
    
    '' Default for Menubar items and
    '' Caption-Separator items is mssClear
    
    '' Default for all other items except for
    '' Sidebar items is mssColor, and for
    '' Sidebars it is mssHotTrack
    
    '' You can override the default styling
    '' with the values, below.
    
    '' Any setting is valid with either a MenuItem object or
    '' Sidebar object
    
    '' 'Or' the value with mssBevel to get a 'beveled' line
    '' effect if the mode draws boxes
    
    mssColor = &H1&
    mssClear = &H2&
    mssHotTrack = &H3&
    
    mssBevel = &H40&
    mssFlat = &HC0&
    
    mssNoCheckBevel = &H100&
End Enum


'' Friend properties that are referenced
'' on measure and draw


Private m_Draw As DRAWITEMSTRUCT
Private m_Measure As MEASUREITEMSTRUCT

Private m_ImageActual As SIZEAPI

Private m_TextRect As RECT
Private m_AccelRect As RECT
Private m_PictureRect As RECT
Private m_ArrowRect As RECT

'' Private variables to hold data for public properties

'' The Item property (that homes this object)
Private m_Item As Object

'' Image Scaling
Private m_ScaleImages As Boolean
Private m_ImageScaleWidth As Long
Private m_ImageScaleHeight As Long

Private m_DrawBoolPic As Boolean

'' Color of item if selected
Private m_SelectForeground As OLE_COLOR
Private m_SelectBackground As OLE_COLOR
Private m_SelectBkGradient As OLE_COLOR

'' Color of item if not selected
Private m_ItemForeground As OLE_COLOR
Private m_ItemBackground As OLE_COLOR
Private m_ItemBkGradient As OLE_COLOR

'' For 'XP-Style' menus
Private m_RunnerColor As OLE_COLOR
Private m_RunnerSelection As OLE_COLOR

'' When a checkmark item is selected/highlighted
Private m_ItemCheckBackground As OLE_COLOR

'' Default Font property
Private WithEvents m_Font As StdFont
Attribute m_Font.VB_VarHelpID = -1

'' Multi-gradient property
Private m_MultiGrad As MultiGradient

'' Multi-gradient property for selection
Private m_SelMultiGrad As MultiGradient

Private m_UserWidth As Long
Private m_UserHeight As Long

Private m_SelectionStyle As SelectionStyleConstants

Private m_PictureVAlign As StdMenuAPI.PictureVAlignConstants
Private m_TextAlign As StdMenuAPI.TextAlignConstants

Private m_WrapLimit As Long

Public Event PropertyChanged(ByVal PropertyName As String)

'' Placement
Public Property Get TextAlign() As TextAlignConstants
    TextAlign = m_TextAlign
End Property

Public Property Let TextAlign(ByVal vData As TextAlignConstants)
    m_TextAlign = vData
    
    RaiseEvent PropertyChanged("TextAlign")
End Property

Public Property Get PictureVAlign() As PictureVAlignConstants
    PictureVAlign = m_PictureVAlign
End Property

Public Property Let PictureVAlign(ByVal vData As PictureVAlignConstants)
    m_PictureVAlign = vData

    RaiseEvent PropertyChanged("PictureVAlign")
End Property

'' Font
Public Property Get Font() As StdFont
    Set Font = m_Font
End Property

Public Property Set Font(ByVal vData As StdFont)
    Set m_Font = vData
    
    If (TypeOf m_Item Is MenuItem) Then
        Set m_Item.Submenu.Font = vData
    End If

    RaiseEvent PropertyChanged("Font")
End Property

'' Colors

Public Property Get SelectBackground() As OLE_COLOR
    SelectBackground = m_SelectBackground
End Property

Public Property Let SelectBackground(ByVal vData As OLE_COLOR)
    m_SelectBackground = vData
    RaiseEvent PropertyChanged("SelectBackground")
End Property

Public Property Get SelectBkGradient() As OLE_COLOR
    SelectBkGradient = m_SelectBkGradient
End Property

Public Property Let SelectBkGradient(ByVal vData As OLE_COLOR)
    m_SelectBkGradient = vData
    RaiseEvent PropertyChanged("SelectBkGradient")
End Property

Public Property Get SelectForeground() As OLE_COLOR
    SelectForeground = m_SelectForeground
End Property

Public Property Let SelectForeground(ByVal vData As OLE_COLOR)
    m_SelectForeground = vData
    RaiseEvent PropertyChanged("SelectForeground")
End Property

Public Property Get ItemBackground() As OLE_COLOR
    ItemBackground = m_ItemBackground
End Property

Public Property Let ItemBackground(ByVal vData As OLE_COLOR)
    m_ItemBackground = vData
    RaiseEvent PropertyChanged("ItemBackground")
End Property

Public Property Get ItemBkGradient() As OLE_COLOR
    ItemBkGradient = m_ItemBkGradient
End Property

Public Property Let ItemBkGradient(ByVal vData As OLE_COLOR)
    m_ItemBkGradient = vData
    RaiseEvent PropertyChanged("ItemBkGradient")
End Property

Public Property Get ItemForeground() As OLE_COLOR
    ItemForeground = m_ItemForeground
End Property

Public Property Let ItemForeground(ByVal vData As OLE_COLOR)
    m_ItemForeground = vData
    RaiseEvent PropertyChanged("ItemForeground")
End Property

Public Property Get ItemCheckBackground() As OLE_COLOR
    ItemCheckBackground = m_ItemCheckBackground
End Property

Public Property Let ItemCheckBackground(ByVal vData As OLE_COLOR)
    m_ItemCheckBackground = vData
    RaiseEvent PropertyChanged("ItemCheckBackground")
End Property

Public Property Get RunnerColor() As OLE_COLOR
    RunnerColor = m_RunnerColor
End Property

Public Property Let RunnerColor(ByVal vData As OLE_COLOR)
    m_RunnerColor = vData
    RaiseEvent PropertyChanged("RunnerColor")
End Property

Public Property Get RunnerSelection() As OLE_COLOR
    RunnerSelection = m_RunnerSelection
End Property

Public Property Let RunnerSelection(ByVal vData As OLE_COLOR)
    m_RunnerSelection = vData
    RaiseEvent PropertyChanged("RunnerSelection")
End Property

Public Property Get MultiGradient() As MultiGradient
    Set MultiGradient = m_MultiGrad
End Property

Public Property Get SelMultiGradient() As MultiGradient
    Set SelMultiGradient = m_SelMultiGrad
End Property

'' If you want to make the item's selection-Style 'clear'
'' at any time, you can set the 'selected' and 'not selected'
'' color schemes to identical values.

'' (when the SelectionStyle property is set to mssClear this function
''  is called)

Public Sub UnifyColorScheme(Optional ByVal fTakeSelection As Boolean)
    Dim varColor As Variant
    
    If (fTakeSelection = True) Then
        Set m_MultiGrad = New MultiGradient
        
        If (m_SelMultiGrad.Count > 0&) Then
            For Each varColor In m_SelMultiGrad
                m_MultiGrad.Add varColor
            Next varColor
        End If
        
        m_ItemForeground = m_SelectForeground
        
        m_ItemBackground = m_SelectBackground
        m_ItemBkGradient = m_SelectBkGradient
        
    Else
        Set m_SelMultiGrad = New MultiGradient
        
        If (m_MultiGrad.Count > 0&) Then
            For Each varColor In m_MultiGrad
                m_SelMultiGrad.Add varColor
            Next varColor
        End If
                
        m_SelectForeground = m_ItemForeground
        
        m_SelectBackground = m_ItemBackground
        m_SelectBkGradient = m_ItemBkGradient
        
    End If
    
End Sub

'' Highlighting

Public Property Get SelectionStyle() As SelectionStyleConstants
    SelectionStyle = m_SelectionStyle
End Property

Public Property Let SelectionStyle(ByVal vData As SelectionStyleConstants)
    If (m_SelectionStyle <> vData) Then
        
        If (vData = mssHotTrack) Then
            If (m_SelectForeground = (&H80000000 Or COLOR_HIGHLIGHTTEXT)) Then
                m_SelectForeground = vbBlue
            End If
            
            If (m_SelectBackground = (&H80000000 Or COLOR_HIGHLIGHT)) Then
                m_SelectBackground = (&H80000000 Or COLOR_MENU)
            End If
        ElseIf (vData = mssDefault) Or (vData = mssColor) Then
            If (m_SelectForeground = vbBlue) Then
                m_SelectForeground = (&H80000000 Or COLOR_HIGHLIGHTTEXT)
            End If
            
            If (m_SelectBackground = (&H80000000 Or COLOR_MENU)) Then
                m_SelectBackground = (&H80000000 Or COLOR_HIGHLIGHT)
            End If
            
        ElseIf (vData And mssFlat) = mssFlat Then
            If (m_SelectForeground = (&H80000000 Or COLOR_HIGHLIGHTTEXT)) Then
                m_SelectForeground = vbBlue
            End If
            
            If (m_SelectBackground = (&H80000000 Or COLOR_HIGHLIGHT)) Then
                m_SelectBackground = (&H80000000 Or COLOR_INFOBK)
            End If
        End If
        
        If (vData = mssClear) Then
            UnifyColorScheme
        End If
        
        m_SelectionStyle = vData
        RaiseEvent PropertyChanged("SelectionStyle")
    End If
        
End Property

Public Property Get DrawBoolPic() As Boolean
    DrawBoolPic = m_DrawBoolPic
End Property

Public Property Let DrawBoolPic(ByVal vData As Boolean)
    m_DrawBoolPic = vData

    RaiseEvent PropertyChanged("DrawBoolPic")
End Property

'' Sizing and Scaling

Public Property Get ScaleImages() As Boolean
    ScaleImages = m_ScaleImages
End Property

Public Property Let ScaleImages(ByVal vData As Boolean)
    m_ScaleImages = vData
    RaiseEvent PropertyChanged("ScaleImages")
End Property

Public Property Get ImageScaleWidth() As Long
    ImageScaleWidth = m_ImageScaleWidth
End Property

Public Property Let ImageScaleWidth(ByVal vData As Long)
    m_ImageScaleWidth = vData

    RaiseEvent PropertyChanged("ImageScaleWidth")
End Property

Public Property Get ImageScaleHeight() As Long
    ImageScaleHeight = m_ImageScaleHeight
End Property

Public Property Let ImageScaleHeight(ByVal vData As Long)
    m_ImageScaleHeight = vData

    RaiseEvent PropertyChanged("ImageScaleHeight")
End Property

Public Sub SizeImage(ByVal ImageWidth, Optional ImageHeight)
    
    m_ImageScaleWidth = ImageWidth
    
    If IsMissing(ImageHeight) = False Then
        m_ImageScaleHeight = ImageHeight
    End If
    
    If (m_ImageScaleWidth = 0&) And (m_ImageScaleHeight = 0&) Then
        m_ScaleImages = False
    Else
        m_ScaleImages = True
    End If
    
    RaiseEvent PropertyChanged("ImageScale")
    
End Sub

Public Property Get UserWidth() As Long
    UserWidth = m_UserWidth
End Property

Public Property Let UserWidth(ByVal vData As Long)
    m_UserWidth = vData
    RaiseEvent PropertyChanged("UserWidth")
End Property

Public Property Get UserHeight() As Long
    UserHeight = m_UserHeight
End Property

Public Property Let UserHeight(ByVal vData As Long)
    m_UserHeight = vData
    RaiseEvent PropertyChanged("UserHeight")
End Property

Public Sub UserSize(Width, Optional Height)
    m_UserWidth = Width
    
    If IsMissing(Height) = False Then
        m_UserHeight = Height
    End If
    
    RaiseEvent PropertyChanged("UserSize")
End Sub

Public Property Get WrapLimit() As Long
    WrapLimit = m_WrapLimit
End Property

Public Property Let WrapLimit(ByVal vData As Long)
    If m_WrapLimit <> vData Then
        m_WrapLimit = vData
        RaiseEvent PropertyChanged("WrapLimit")
    End If
End Property

'' Item Property

Public Property Get Item() As Object
    Set Item = m_Item
End Property

Friend Property Set Item(ByVal vData As Object)
    Set m_Item = vData
End Property

Public Sub LoadProperties(PropBag As PropertyBag, Optional Prefix As String)
    Dim sPrefix As String
    
    If (Prefix <> "") Then sPrefix = Prefix + "."
    
    Set m_Font = PropBag.ReadProperty(sPrefix + "Font", m_Font)
    
    m_MultiGrad.LoadProperties PropBag, sPrefix + "MultiGrad"
    m_SelMultiGrad.LoadProperties PropBag, sPrefix + "SelMultiGrad"
    
    m_ItemForeground = PropBag.ReadProperty(sPrefix + "ItemForeground", m_ItemForeground)
    m_ItemBackground = PropBag.ReadProperty(sPrefix + "ItemBackground", m_ItemBackground)
    m_ItemBkGradient = PropBag.ReadProperty(sPrefix + "ItemBkGradient", m_ItemBkGradient)
    
    m_RunnerColor = PropBag.ReadProperty(sPrefix + "RunnerColor", m_RunnerColor)
    m_RunnerSelection = PropBag.ReadProperty(sPrefix + "RunnerSelection", m_RunnerSelection)
    
    m_ItemCheckBackground = PropBag.ReadProperty(sPrefix + "ItemCheckBackground", m_ItemCheckBackground)
    
    m_SelectForeground = PropBag.ReadProperty(sPrefix + "SelectForeground", m_SelectForeground)
    m_SelectBackground = PropBag.ReadProperty(sPrefix + "SelectBackground", m_SelectBackground)
    m_SelectBkGradient = PropBag.ReadProperty(sPrefix + "SelectBkGradient", m_SelectBkGradient)
    
    m_UserWidth = PropBag.ReadProperty(sPrefix + "UserWidth", m_UserWidth)
    m_UserHeight = PropBag.ReadProperty(sPrefix + "UserHeight", m_UserHeight)
    
    m_ImageScaleWidth = PropBag.ReadProperty(sPrefix + "ImageScaleWidth", m_ImageScaleWidth)
    m_ImageScaleHeight = PropBag.ReadProperty(sPrefix + "ImageScaleHeight", m_ImageScaleHeight)
    
    m_ScaleImages = PropBag.ReadProperty(sPrefix + "ScaleImages", m_ScaleImages)

    m_SelectionStyle = PropBag.ReadProperty(sPrefix + "SelectionStyle", m_SelectionStyle)
    m_DrawBoolPic = PropBag.ReadProperty(sPrefix + "DrawBoolPic", m_DrawBoolPic)

    m_TextAlign = PropBag.ReadProperty(sPrefix + "TextAlign", m_TextAlign)
    m_PictureVAlign = PropBag.ReadProperty(sPrefix + "PictureVAlign", m_PictureVAlign)

End Sub

Public Sub SaveProperties(PropBag As PropertyBag, Optional Prefix As String)
    Dim sPrefix As String
    
    If (Prefix <> "") Then sPrefix = Prefix + "."
    
    PropBag.WriteProperty sPrefix + "Font", m_Font
    m_MultiGrad.SaveProperties PropBag, sPrefix + "MultiGrad"
    m_SelMultiGrad.SaveProperties PropBag, sPrefix + "SelMultiGrad"
    
    PropBag.WriteProperty sPrefix + "ItemForeground", m_ItemForeground
    PropBag.WriteProperty sPrefix + "ItemBackground", m_ItemBackground
    PropBag.WriteProperty sPrefix + "ItemBkGradient", m_ItemBkGradient
        
    PropBag.WriteProperty sPrefix + "RunnerColor", m_RunnerColor
    PropBag.WriteProperty sPrefix + "RunnerSelection", m_RunnerSelection
        
    PropBag.WriteProperty sPrefix + "ItemCheckBackground", m_ItemCheckBackground
    
    PropBag.WriteProperty sPrefix + "SelectForeground", m_SelectForeground
    PropBag.WriteProperty sPrefix + "SelectBackground", m_SelectBackground
    PropBag.WriteProperty sPrefix + "SelectBkGradient", m_SelectBkGradient
        
    PropBag.WriteProperty sPrefix + "UserWidth", m_UserWidth
    PropBag.WriteProperty sPrefix + "UserHeight", m_UserHeight
        
    PropBag.WriteProperty sPrefix + "ImageScaleWidth", m_ImageScaleWidth
    PropBag.WriteProperty sPrefix + "ImageScaleHeight", m_ImageScaleHeight
        
    PropBag.WriteProperty sPrefix + "ScaleImages", m_ScaleImages
    
    PropBag.WriteProperty sPrefix + "SelectionStyle", m_SelectionStyle
    PropBag.WriteProperty sPrefix + "DrawBoolPic", m_DrawBoolPic
    
    PropBag.WriteProperty sPrefix + "TextAlign", m_TextAlign
    PropBag.WriteProperty sPrefix + "PictureVAlign", m_PictureVAlign
    
End Sub

Friend Sub SetDraw(lpInfo As DRAWITEMSTRUCT)
    CopyMemory m_Draw, lpInfo, Len(m_Draw)
End Sub

Friend Sub SetMeasure(lpInfo As MEASUREITEMSTRUCT)
    CopyMemory m_Measure, lpInfo, Len(m_Measure)
End Sub

Friend Sub GetDraw(lpInfo As DRAWITEMSTRUCT)
    CopyMemory lpInfo, m_Draw, Len(m_Draw)
End Sub

Friend Sub GetMeasure(lpInfo As MEASUREITEMSTRUCT)
    CopyMemory lpInfo, m_Measure, Len(m_Measure)
End Sub


'' Private Events

Private Sub m_Font_FontChanged(ByVal PropertyName As String)
    On Error Resume Next
    
    If TypeOf m_Item.Parent Is Menubar Then
        DrawMenuBar m_Item.Parent.hWnd
    End If
        
    RaiseEvent PropertyChanged("Font." + PropertyName)
    
End Sub

Private Sub Class_Initialize()
    Dim cr1 As Long, _
        cr2 As Long
    
    
    m_TextAlign = taLeft
    m_PictureVAlign = ppCenter
    
    Set m_MultiGrad = New MultiGradient
    Set m_SelMultiGrad = New MultiGradient
    
    m_ItemForeground = (&H80000000 Or COLOR_MENUTEXT)
    m_ItemBackground = (&H80000000 Or COLOR_MENU)
    m_ItemBkGradient = -1&
    
    m_ItemCheckBackground = &H80000000 Or COLOR_APPWORKSPACE
    
    m_SelectForeground = &H80000000 Or COLOR_HIGHLIGHTTEXT
    m_SelectBackground = &H80000000 Or COLOR_HIGHLIGHT
    m_SelectBkGradient = -1&

    cr1 = GetSysColor(COLOR_HOTLIGHT)
    cr2 = &HBBFFFF
        
    m_RunnerColor = GetAverageColor(cr2, cr1)
    cr2 = &HAFAFAF
        
    m_RunnerSelection = GetAverageColor(cr2, cr1)
    
    m_ScaleImages = True
    SizeImage 16, 16

End Sub

Private Sub Class_Terminate()
    Set m_MultiGrad = Nothing
    Set m_SelMultiGrad = Nothing
End Sub

Private Sub Class_ReadProperties(PropBag As PropertyBag)
    
    LoadProperties PropBag
          
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)

    SaveProperties PropBag

End Sub

