VERSION 5.00
Begin VB.UserControl MenuPreview 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2835
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3810
   ControlContainer=   -1  'True
   ScaleHeight     =   189
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   254
End
Attribute VB_Name = "MenuPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' GUInerd Standard Menu System
''' Version 4.0&

''' Objects/API Dll

''' Menu Designer Active-X Control


''' ****************************************************************************
''' *** Check the end of the file for copyright and distribution information ***
''' ****************************************************************************

Option Explicit

Private Editing As Boolean
Private EditPic As Boolean
Private Sizing As Boolean
Private UI_Init As Boolean

Private PaintDC As Long
Private PaintBitmap As Long
Private PaintNoResize As Boolean

Private WithEvents EditBox As EditWnd
Attribute EditBox.VB_VarHelpID = -1
Private WithEvents Scroll As WndScroll
Attribute Scroll.VB_VarHelpID = -1
Private WithEvents ThreadMouse As MouseHook
Attribute ThreadMouse.VB_VarHelpID = -1

Private EditItem_Orig As MenuItem
Private EditItem As MenuItem

Private WithEvents EditSizer As SizeRect
Attribute EditSizer.VB_VarHelpID = -1

' Private WithEvents EditBox As EditWnd

Public Enum MenuPreviewAlignment
    mpaCenter = 0&
    mpaLeft = 1
    mpaRight = 2
    mpaTop = 3
    mpaBottom = 4
    mpaLeftTop = 5
    mpaLeftBottom = 6
    mpaRightTop = 7
    mpaRightBottom = 8
End Enum

Public Enum ScrollbarConstants
    sbNone = 0&
    sbHorizontal = 1
    sbVertical = 2
    sbBoth = 3
End Enum

Public Enum EditModeConstants
    emAuto = 0&
    emManual = 1&
End Enum

Public Enum ItemHitTestPlaceConstants
    htText = 0&
    htPicture = 1&
    htSubmenu = 2&
    htOther = 3&
End Enum

Private m_Col As Collection

Private m_Current As MenuItem

Private WithEvents m_Font As StdFont
Attribute m_Font.VB_VarHelpID = -1

Private m_MenuType As MenuTypeConstants
Private m_EditMode As EditModeConstants
Private m_Alignment As MenuPreviewAlignment
Private m_Scrollbars As StandardMenu_ObjectsAPI.ScrollbarConstants

Private m_BackColor As OLE_COLOR

Private m_HScrollSpace As Long
Private m_VScrollSpace As Long

Private m_RightToLeft As Boolean
Private m_Events As Boolean
Private m_DisableNoScroll As Boolean
Private m_ScrollTrackPercent As Boolean

Public Event SubmenuSelect(Item As MenuItem)
Public Event ItemSelect(Item As MenuItem)
Public Event BeforeItemEdit()
Public Event AfterItemEdit()

Public Property Get ScrollTrackPercent() As Boolean
    ScrollTrackPercent = m_ScrollTrackPercent
End Property

Public Property Let ScrollTrackPercent(ByVal vData As Boolean)
    m_ScrollTrackPercent = vData
    PropertyChanged "ScrollTrackPercent"
    
    UserControl_Paint
End Property

Public Function Add(ByVal Caption As String, Optional ByVal Key As String, Optional ByVal Picture As StdPicture, Optional ByVal InsertAfter As Integer = -1, Optional ByVal CopyItem As MenuItem) As MenuItem

    Dim varNewItem As New MenuItem
    
    varNewItem.Caption = Caption
    
    varNewItem.Key = Key
    Set varNewItem.Picture = Picture
    Set varNewItem.Parent = Me
    
    varNewItem.ParentType = m_MenuType
    
    If (Key <> "") And (InsertAfter <> -1) Then
        m_Col.Add varNewItem, Key, , InsertAfter
    ElseIf (InsertAfter <> -1) Then
        m_Col.Add varNewItem, , , InsertAfter
    ElseIf (Key <> "") Then
        m_Col.Add varNewItem, Key
    Else
        m_Col.Add varNewItem
    End If

    Set Add = varNewItem
    Set varNewItem = Nothing

    If Not CopyItem Is Nothing Then
        varNewItem.CopyItem CopyItem, True, True
    End If
    
    UserControl_Paint

End Function

Public Sub Remove(Index)
    On Error Resume Next
    
    m_Col.Remove Index
    
    DeleteDC PaintDC
    PaintDC = 0&
    
    DeleteObject PaintBitmap
    PaintBitmap = 0&
    
    UserControl_Paint

End Sub

Public Sub Clear()
        
    Set m_Col = Nothing
    Set m_Col = New Collection
    
    DeleteDC PaintDC
    PaintDC = 0&
    
    DeleteObject PaintBitmap
    PaintBitmap = 0&
    
    UserControl_Paint

End Sub

Public Property Get Item(Index) As MenuItem

    On Error Resume Next
    Set Item = m_Col(Index)
    
End Property

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_Col.[_NewEnum]
End Function

Public Property Get Font() As StdFont
    Set Font = m_Font
End Property

Public Property Set Font(ByVal vData As StdFont)
    Set m_Font = vData
    PropertyChanged "Font"
End Property

Public Property Get Alignment() As MenuPreviewAlignment
    Alignment = m_Alignment
End Property

Public Property Let Alignment(ByVal vData As MenuPreviewAlignment)
    m_Alignment = vData
    PropertyChanged "Alignment"
    
    DeleteDC PaintDC
    PaintDC = 0&
    
    DeleteObject PaintBitmap
    PaintBitmap = 0&
    
    UserControl_Paint
End Property

Public Property Get Scrollbars() As StandardMenu_ObjectsAPI.ScrollbarConstants
    Scrollbars = m_Scrollbars
End Property

Public Property Let Scrollbars(ByVal vData As StandardMenu_ObjectsAPI.ScrollbarConstants)
    m_Scrollbars = vData
    PropertyChanged "ScrollBars"
    
    Scroll.Scrollbars = m_Scrollbars
    
    DeleteDC PaintDC
    PaintDC = 0&
    
    DeleteObject PaintBitmap
    PaintBitmap = 0&
    
    UserControl_Paint
End Property

Public Property Get DisableNoScroll() As Boolean
    DisableNoScroll = m_DisableNoScroll
End Property

Public Property Let DisableNoScroll(ByVal vData As Boolean)
    m_DisableNoScroll = vData
    
    Scroll.DisableNoScroll = m_DisableNoScroll
    
    PropertyChanged "DisableNoScroll"
End Property

Public Property Get Events() As Boolean
    Events = m_Events
End Property

Public Property Let Events(ByVal vData As Boolean)
    m_Events = vData
    PropertyChanged "Events"
End Property

Public Property Get EditMode() As EditModeConstants
    EditMode = m_EditMode
End Property

Public Property Let EditMode(ByVal vData As EditModeConstants)
    m_EditMode = vData
    PropertyChanged "EditMode"
End Property

Public Property Get RightToLeft() As Boolean
    RightToLeft = m_RightToLeft
End Property

Public Property Let RightToLeft(ByVal vData As Boolean)
    m_RightToLeft = vData
    UserControl.RightToLeft = vData
    
    PropertyChanged "RightToLeft"
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal vData As OLE_COLOR)
    m_BackColor = vData
    UserControl.BackColor = vData
    
    PropertyChanged "BackColor"
    
End Property

Public Property Get MenuType() As MenuTypeConstants
    MenuType = m_MenuType
End Property

Public Property Let MenuType(ByVal vData As MenuTypeConstants)
    
    If (m_MenuType <> vData) Then
        m_MenuType = vData
        PropertyChanged "MenuType"
    
        Dim varObj As MenuItem
    
        For Each varObj In m_Col
            varObj.ParentType = vData
        Next varObj
                    
        If (m_MenuType = mtcMenubar) Then
            
            Select Case m_Alignment
            
                Case mpaCenter, mpaTop, mpaLeftTop, mpaRightTop, mpaLeft
                
                    m_Alignment = mpaTop
                    
                Case Else
                    
                    m_Alignment = mpaBottom
                    
            End Select
            
        End If
        
        UserControl_Paint
        
    End If
        
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property



Public Sub BeginItemEdit(Optional ByVal Item As MenuItem)
        
    If (Editing = True) Or (Sizing = True) Then Exit Sub
    
    If Item Is Nothing Then
        Set Item = EditItem
    Else
        Set EditItem = Item
    End If
    
    If Item Is Nothing Then Exit Sub
    
    SelectItem
    
    EditBox.BackColor = Item.ItemBackground
    EditBox.TextColor = Item.ItemForeground
    EditBox.Text = Item.Caption
    
    Set EditItem_Orig = New MenuItem
    EditItem_Orig.CopyItem Item
    
    Editing = True
    SizeEdit Item

    EditBox.Visible = True
    EditBox.Enabled = True
    
End Sub

Public Sub BeginItemResize(ByVal Item As MenuItem)
        
    If (Editing = True) Or (Sizing = True) Then Exit Sub
    
    SelectItem
    SetSizeBox Item
    EditSizer.Attach hWnd, PaintDC
    EditSizer.Visible = True
    
    Set EditItem = Item
    Sizing = True
    
End Sub

Public Sub EndItemEdit()
    
End Sub

Public Sub CancelItemEdit()

End Sub

Public Sub Redraw()
    UserControl_Paint
End Sub

Public Function HitTest(ByVal x As Long, ByVal y As Long, Optional cbPlace As ItemHitTestPlaceConstants) As MenuItem

    Dim varObj As MenuItem, _
        m_Draw As DRAWITEMSTRUCT
        
    Dim rx As Long, _
        ry As Long
        
    Dim zx As Long, _
        zy As Long
    
    Dim lpRect As RECT
    
    rx = x + Scroll.HPosition
    ry = y + Scroll.VPosition
    
    For Each varObj In m_Col
    
        varObj.GetLastDraw m_Draw
        
        If WithinRect(m_Draw.rcItem, rx, ry) = True Then
            
            Set HitTest = varObj
               
            If IsMissing(cbPlace) = False Then
                
                cbPlace = htOther
                
                varObj.GetTextRect lpRect
                
                If WithinRect(lpRect, rx, ry) = True Then
                    cbPlace = htText
                End If
            
                varObj.GetPictureRect lpRect
                
                If WithinRect(lpRect, rx, ry) = True Then
                    cbPlace = htPicture
                End If
                
                varObj.GetDropRect lpRect
                                
                If WithinRect(lpRect, rx, ry) = True Then
                    cbPlace = htSubmenu
                End If
                
            End If
            
            Exit Function
        
        End If
    
    Next varObj
    
End Function

Friend Sub MenuSelectEvent(ByVal Item As MenuItem)
    RaiseEvent ItemSelect(Item)
End Sub

Private Sub StopEdit(Optional ByVal Cancel As Boolean)
        
    EditBox.SetRedraw False
    EditBox.Visible = False
    EditBox.Enabled = False
    
    Editing = False
        
    If (Cancel = True) Then
        EditItem.Caption = EditItem_Orig.Caption
    Else
        EditItem.Caption = EditBox.Text
    End If
        
    Set EditItem = Nothing
    Set EditItem_Orig = Nothing
    
    UserControl_Paint

End Sub

Private Sub StopResize()
    
    If (Sizing = True) Then
        Sizing = False
        
        EditSizer.Visible = False
        EditSizer.Detach
        
        Set EditItem = Nothing
        Set EditItem_Orig = Nothing
        
        UserControl_Paint
    End If
        
End Sub

Private Sub SetSizeBox(Optional ByVal Item As MenuItem)
    Dim lpDraw As DRAWITEMSTRUCT, _
        cx As Long, _
        cy As Long
        
    If Item Is Nothing Then Set Item = EditItem
    If Item Is Nothing Then Exit Sub
    
    Item.GetLastDraw lpDraw
    
    With lpDraw.rcItem
        cx = (.Right - .Left) + 1
        cy = (.Bottom - .Top) + 1
    
        EditSizer.Move .Left, .Top, cx, cy
    End With
    
End Sub

Private Sub SizeEdit(ByVal Item As MenuItem)
    Dim lpRect As RECT, _
        rcWnd As RECT
    
    Dim cx As Long, _
        cy As Long
    
    Item.GetTextRect lpRect
    
    With lpRect
        .Left = .Left - (Scroll.HPosition)
        .Top = .Top - (Scroll.VPosition)
        .Right = .Right - (Scroll.HPosition)
        .Bottom = .Bottom - (Scroll.VPosition)
    End With
    
    cx = (lpRect.Right - lpRect.Left) - 2
    cy = (lpRect.Bottom - lpRect.Top) - 2
    
    If (Item.WrapLimit >= WrapLimit_Minimum) And (Item.WrapLimit < cx) Then cx = Item.WrapLimit
    
    With EditBox
    
        If ((.Left <> lpRect.Left) Or (.Top <> (lpRect.Top + 2)) Or _
            (.Width <> cx) Or (.Height <> cy)) Then
    
            EditBox.Move lpRect.Left, lpRect.Top + 2, cx, cy
        End If
    
    End With
    
End Sub

Private Sub UpdateView(Optional ByVal cx As Long, Optional ByVal cy As Long, Optional vItem As MenuItem)

    Dim x As Long, _
        y As Long
        
    Dim lpSize As SIZEAPI
            
    Dim iWidth As Long, _
        iHeight As Long
        
    Dim lpRect As RECT, _
        lpDraw As DRAWITEMSTRUCT
    
    Dim rcItem As RECT
        
    If (IsMissing(cx) = False) And (cx <> 0&) Then
        Scroll.HPosition = Scroll.HPosition + cx
    End If
    
    If (IsMissing(cy) = False) And (cy <> 0&) Then
        Scroll.VPosition = Scroll.VPosition + cy
    End If
    
    iWidth = m_HScrollSpace + ScaleWidth
    iHeight = m_VScrollSpace + ScaleHeight
    
    x = Scroll.HPosition
    y = Scroll.VPosition
            
    If (x + ScaleWidth) > iWidth Then
        x = iWidth - ScaleWidth
        Scroll.HPosition = x
        x = Scroll.HPosition
    End If
                
    If (y + ScaleHeight) > iHeight Then
        y = iHeight - ScaleHeight
        Scroll.VPosition = y
        y = Scroll.VPosition
    End If
    
    lpSize.cx = (iWidth - x)
    lpSize.cy = (iHeight - y)
            
    If Not vItem Is Nothing Then
        vItem.GetLastDraw lpDraw
        CopyMemory rcItem, lpDraw.rcItem, Len(rcItem)
                
        lpRect.Left = rcItem.Left - x
        lpRect.Top = rcItem.Top - y
        
        x = rcItem.Left
        y = rcItem.Top
        
        lpSize.cx = (rcItem.Right - rcItem.Left) + 1
        lpSize.cy = (rcItem.Bottom - rcItem.Top) + 1
    
        BitBlt hDC, lpRect.Left, lpRect.Top, lpSize.cx, lpSize.cy, PaintDC, x, y, SRCCOPY
    Else
        BitBlt hDC, 0&, 0&, lpSize.cx, lpSize.cy, PaintDC, x, y, SRCCOPY
    End If
    
End Sub

Private Function ConfigureScrollSpace() As Long
    Dim i As Long, _
        oX As Long, _
        oY As Long
    
    Dim lpRect As RECT, _
        hBrush As Long, _
        lpBrush As LOGBRUSH
    
    Dim Col() As SIZEAPI, _
        StartPos As POINTAPI, _
        lpBands() As Long
    
    Dim tWidth As Long, _
        tHeight As Long
    
    Dim pX As Single, _
        pY As Single
        
    Dim oHorz As Long, _
        oVert As Long
    
    Dim mhTotal As Long, _
        mvTotal As Long
    
    Dim ChangeDC As Boolean
    
    Dim varNCLI As NONCLIENTMETRICS
    
    varNCLI.cbSize = Len(varNCLI)
    SystemParametersInfo SPI_GETNONCLIENTMETRICS, Len(varNCLI), varNCLI, 0&
    
    oHorz = m_HScrollSpace
    oVert = m_VScrollSpace
        
    oX = Scroll.HPosition
    oY = Scroll.VPosition
    
    If (oX <> 0&) And (m_HScrollSpace <> 0&) Then
        pX = oX / (m_HScrollSpace)
    End If
    
    If (oY <> 0&) And (m_VScrollSpace <> 0&) Then
        pY = oY / (m_VScrollSpace)
    End If
            
    i = MeasureItems
    ConfigureScrollSpace = i
    
    m_HScrollSpace = 0&
    m_VScrollSpace = 0&
    
    If (i <= 0&) Then Exit Function
    
    If (m_MenuType = mtcPopup) Then
        
        PopupMenuRect lpRect, Col()
                
        mhTotal = (lpRect.Right + 24)
        mvTotal = (lpRect.Bottom + 24)
        
        tWidth = mhTotal
        tHeight = mvTotal
        
    Else
    
        MeasureMenubar lpRect, lpBands
        
        mhTotal = lpRect.Right
        mvTotal = lpRect.Bottom
        
        tWidth = mhTotal
        tHeight = mvTotal
        
    End If
    
    If (tWidth < ScaleWidth) Then tWidth = ScaleWidth
    If (tHeight < ScaleHeight) Then tHeight = ScaleHeight
    
    '' Check to see if the menu scroll bars will change
    '' if so, resize the DC to fit the total drawing area.
    
    If (Scroll.DisableNoScroll = False) Then
        
        If (oVert <> 0&) And (tHeight <= ScaleHeight) Then
            ChangeDC = True
        End If
        
        If (oHorz <> 0&) And (tWidth <= ScaleWidth) Then
            ChangeDC = True
        End If
                            
    End If
    
    m_HScrollSpace = tWidth - ScaleWidth
    m_VScrollSpace = tHeight - ScaleHeight
                                
    If (Scroll.DisableNoScroll = False) Then
        
        If (m_HScrollSpace <> 0&) And (oHorz = 0&) Then
            If (m_HScrollSpace <= varNCLI.iScrollWidth) And _
                (m_VScrollSpace <= 0&) Then
                
                m_HScrollSpace = 0&
            End If
            
            ChangeDC = True
        End If
            
        If (m_VScrollSpace <> 0&) And (oVert = 0&) Then
            If (m_VScrollSpace <= varNCLI.iScrollHeight) And _
                (m_HScrollSpace <= 0&) Then
                
                m_VScrollSpace = 0&
            End If
            
            ChangeDC = True
        End If
        
        If (m_HScrollSpace <> oHorz) Or _
            (m_VScrollSpace <> oVert) Then ChangeDC = True
    End If
    
    If (ChangeDC = True) Then
        If (PaintBitmap <> 0&) Then DeleteObject PaintBitmap
        
        PaintBitmap = 0&
    End If
    
    If (m_HScrollSpace <> 0&) Then
        Scroll.SetRange sbHorizontal, 0&, tWidth
        Scroll.HPage = ScaleWidth
    Else
        Scroll.SetRange sbHorizontal, 0&, 0&
    End If
    
    If (oX <= m_HScrollSpace) Then
        If (m_ScrollTrackPercent = True) Then
            Scroll.HPosition = tWidth * pX
        Else
            Scroll.HPosition = oX
        End If
    Else
        Scroll.HPosition = m_HScrollSpace
    End If
    
    If (m_VScrollSpace <> 0&) Then
        Scroll.SetRange sbVertical, 0&, tHeight
        Scroll.VPage = ScaleHeight
    Else
        Scroll.SetRange sbVertical, 0&, 0&
    End If
    
    If (oY <= m_VScrollSpace) Then
        If (m_ScrollTrackPercent = True) Then
            Scroll.VPosition = tHeight * pY
        Else
            Scroll.VPosition = oY
        End If
    Else
        Scroll.VPosition = m_VScrollSpace
    End If

End Function

Private Sub DrawEditBox(rcBox As RECT, Optional ByVal fgColor As Long = -1&, Optional ByVal bgColor As Long = -1&)
    
    Dim lpRect As RECT, _
        hPen As Long, _
        hOldPen As Long, _
        crColor As Long, _
        oldBgColor As Long
        
    Dim lpPoint As POINTAPI
    
    If (rcBox.Right - rcBox.Left) = 0& Then
        lpRect.Left = (EditBox.Left + Scroll.HPosition) - 3
        lpRect.Top = (EditBox.Top + Scroll.VPosition) - 3
        lpRect.Right = (lpRect.Left + EditBox.Width) + 2
        lpRect.Bottom = (lpRect.Top + EditBox.Height) + 2
    
        EditSizer.Move lpRect.Left, lpRect.Top, (lpRect.Right - lpRect.Left), (lpRect.Bottom - lpRect.Top)
    Else
        CopyMemory lpRect, rcBox, 16&
    End If
    
    If (fgColor = -1&) Then
        crColor = GetSysColor(COLOR_HIGHLIGHT)
    Else
        crColor = fgColor
    End If
    
    If (bgColor <> -1&) Then
        oldBgColor = SetBkColor(PaintDC, bgColor)
    End If
    
    hPen = CreatePen(PS_DOT, 1&, crColor)
    
    hOldPen = SelectObject(PaintDC, hPen)
    
    With lpRect
        MoveToEx PaintDC, .Left, .Top, lpPoint
        LineTo PaintDC, .Right, .Top
        LineTo PaintDC, .Right, .Bottom
        LineTo PaintDC, .Left, .Bottom
        LineTo PaintDC, .Left, .Top
        
        .Left = .Left - 1
        .Top = .Top - 1
        
        .Right = .Right + 1
        .Bottom = .Bottom + 1
        
        MoveToEx PaintDC, .Left, .Top, lpPoint
        LineTo PaintDC, .Right, .Top
        LineTo PaintDC, .Right, .Bottom
        LineTo PaintDC, .Left, .Bottom
        LineTo PaintDC, .Left, .Top
    
    End With
    
    SelectObject PaintDC, hOldPen
    DeleteObject hPen
    
    If (bgColor <> -1&) Then
        SetBkColor PaintDC, oldBgColor
    End If
    
End Sub


Private Function MeasureItems() As Long
    Dim varObj As MenuItem, _
        i As Long
        
    For Each varObj In m_Col
    
        varObj.Measure hWnd, PaintDC
        i = i + 1
    
    Next varObj
    
    MeasureItems = i
    
End Function

Private Sub GetStartPos(lpPoint As POINTAPI, ByVal cx As Long, ByVal cy As Long, ByVal WndWidth As Long, ByVal WndHeight As Long)

    Dim tWidth As Long, _
        tHeight As Long
        
    tWidth = WndWidth
    tHeight = WndHeight

    With lpPoint
        Select Case m_Alignment
    
            Case mpaLeftTop, mpaTop, mpaRightTop
            
                .y = 12
                
            Case mpaBottom, mpaRightBottom, mpaLeftBottom
                        
                .y = tHeight - (12 + cy)
            
            Case mpaCenter, mpaLeft, mpaRight
            
                .y = (tHeight / 2) - (cy / 2)
            
        End Select
        
        Select Case m_Alignment
        
            Case mpaLeftTop, mpaLeft, mpaLeftBottom:
            
                .x = 12
                
            Case mpaBottom, mpaCenter, mpaTop:
            
                .x = (tWidth / 2) - (cx / 2)
            
            Case mpaRightTop, mpaRight, mpaRightBottom:
            
                .x = tWidth - (12 + cx)
        
        End Select
                        
    End With
    
End Sub

Private Function MeasureMenubar(lpRect As RECT, lpBands() As Long) As Long

    Dim x As Long, _
        y As Long
                
    Dim bx As Long
    
    Dim b As Long, _
        i As Long
    
    Dim hBands() As Long
    
    Dim varObj As MenuItem
                
    ReDim hBands(0&)
    For Each varObj In m_Col
    
        If (varObj.Separator = True) And _
            (varObj.SeparatorType And (mstBreak + mstBarBreak)) Then
                
            If (y <= 0&) Then
                y = varObj.MeasureHeight
            End If
        
            hBands(b) = y + 5
            y = varObj.MeasureHeight
            
            b = b + 1
            ReDim Preserve hBands(0& To b)
            
            If (bx < x) Then bx = x
            x = 16 + varObj.MeasureWidth
        Else
            If (y < varObj.MeasureHeight) Then
                y = varObj.MeasureHeight
            End If
    
            x = x + 16 + varObj.MeasureWidth
        End If
    
    Next varObj
    
    hBands(b) = y + 5
    
    If (bx < x) Then bx = x
    
    x = bx
    
    lpRect.Left = 0&
    lpRect.Top = 0&
    
    y = 0&
    For i = 0& To b
        y = y + (5 + hBands(i))
    Next i
    
    lpRect.Bottom = y
    
    If x < (ScaleWidth - 1) Then
        lpRect.Right = ScaleWidth
    Else
        lpRect.Right = x
    End If
    
    lpBands = hBands
    MeasureMenubar = b
            
End Function

Private Sub DrawMenuBar()
    
    Dim varObj As MenuItem
        
    Dim lpPoint As POINTAPI, _
        ItemPts As POINTAPI
    
    Dim lpRect As RECT, _
         dRect As RECT
    
    Dim hBrush As Long, _
        lpBrush As LOGBRUSH
        
    Dim hPen As Long
    
    Dim x As Long, _
        y As Long, _
        pY As Long
        
    Dim Cols As Long, _
        lpBands() As Long, _
        b As Long
    
    Dim pA As POINTAPI, pB As POINTAPI
    
    Cols = MeasureMenubar(lpRect, lpBands)

    Select Case m_Alignment
    
        Case mpaBottom, mpaRightBottom, mpaLeftBottom, mpaRight
            pY = (ScaleHeight + m_VScrollSpace) - lpRect.Bottom
        
        Case Else
            pY = 0&
        
    End Select
    
    With lpRect
        .Top = .Top + pY
        .Bottom = .Bottom + pY
    End With
    
    DrawWindowFrame PaintDC, lpRect
    
    x = 2
    y = pY + 2

    dRect.Left = x
    dRect.Top = y
    
    For Each varObj In m_Col
    
        If (varObj.Separator = True) Then
            If (varObj.SeparatorType = mstBarBreak) Then
            
                pA.x = 2
                pA.y = dRect.Bottom + 2
                
                pB.x = (lpRect.Right)
                pB.y = pA.y
                
                BevelLine PaintDC, pA, pB
            End If
            
            If (varObj.SeparatorType And (mstBreak + mstBarBreak)) Then
            
                dRect.Left = 2
                dRect.Top = dRect.Bottom + 5
                b = b + 1
            
            End If
        
        End If
        
        dRect.Right = dRect.Left + varObj.MeasureWidth + 12
        dRect.Bottom = dRect.Top + lpBands(b)
        
        varObj.Draw hWnd, PaintDC, dRect, pvsNot
        
        dRect.Left = dRect.Right
        
    Next varObj

End Sub

Private Sub PopupMenuRect(lpRect As RECT, Columns() As SIZEAPI)
        
    Dim varObj As MenuItem
    
    Dim c As Long, _
        i As Long, _
        cx As Long, _
        cy As Long
    
    Dim Sizes() As SIZEAPI
    
    ReDim Sizes(0&)
    
    Sizes(0&).cx = 6
    Sizes(0&).cy = 12
    
    For Each varObj In m_Col
    
        If (varObj.Separator = True) Then
            If (varObj.SeparatorType And (mstBarBreak + mstBreak)) Then
            
                Sizes(c).cx = cx
                Sizes(c).cy = cy
                
                c = c + 1
                ReDim Preserve Sizes(0& To c)
                cx = 0&
                cy = 0&
                            
            End If
        End If
        
        If (cx < varObj.MeasureWidth) Then
            cx = varObj.MeasureWidth + 2
        End If
        
        cy = cy + varObj.MeasureHeight + 4
        
    Next varObj
    
    Sizes(c).cx = cx
    Sizes(c).cy = cy
    
    cy = 0&
    cx = 0&
    
    For i = 0& To c
    
        If ((Sizes(i).cy + 4) > cy) Then
            cy = Sizes(i).cy + 4
        End If
        
        cx = cx + Sizes(i).cx
            
    Next i
        
    lpRect.Top = 0&
    lpRect.Left = 0&
    lpRect.Right = cx + (5 * (c + 1))
    lpRect.Bottom = cy + 4
        
    Columns = Sizes
        
End Sub

Private Sub DrawPopupMenu()
    
    Dim varObj As MenuItem
        
    Dim lpPoint As POINTAPI, _
        ItemPts As POINTAPI
    
    Dim lpRect As RECT, _
         dRect As RECT
    
    Dim hBrush As Long, _
        lpBrush As LOGBRUSH
        
    Dim hPen As Long
    
    Dim x As Long, _
        y As Long
    
    Dim Cols() As SIZEAPI, _
        i As Long, _
        c As Long
    
    Dim pA As POINTAPI, pB As POINTAPI
    
    PopupMenuRect lpRect, Cols
    GetStartPos lpPoint, lpRect.Right, lpRect.Bottom, (ScaleWidth + m_HScrollSpace), (ScaleHeight + m_VScrollSpace)
    
    With lpPoint
        lpRect.Left = lpRect.Left + .x
        lpRect.Right = lpRect.Right + .x
        
        lpRect.Top = lpRect.Top + .y
        lpRect.Bottom = lpRect.Bottom + .y
    End With
    
    DrawWindowFrame PaintDC, lpRect
            
    dRect.Left = lpRect.Left + 3
    dRect.Top = lpRect.Top + 3
    dRect.Right = lpRect.Left + Cols(0&).cx
    
    For Each varObj In m_Col
    
        If (varObj.Separator = True) Then
            If (varObj.SeparatorType And mstBarBreak) Then
            
                pA.x = dRect.Left + Cols(c).cx
                pA.y = lpRect.Top + 2
                
                pB.x = pA.x
                pB.y = lpRect.Bottom - 2
                                
                BevelLine PaintDC, pA, pB, True, True
            End If
            
            If (varObj.SeparatorType And (mstBreak + mstBarBreak)) Then
            
                dRect.Left = dRect.Left + Cols(c).cx + 4
                dRect.Top = lpRect.Top + 3
                
                c = c + 1
            
                dRect.Right = Cols(c).cx + dRect.Left
            End If
        
        End If
        
        dRect.Bottom = dRect.Top + varObj.MeasureHeight + 3
        
        varObj.Draw hWnd, PaintDC, dRect, pvsNot
        
        dRect.Top = dRect.Bottom + 1
        
    Next varObj

End Sub

Private Sub SelectItem(Optional ByVal objItem As MenuItem)
    Dim lpRect As RECT
    
    If (Not m_Current Is Nothing) Then
        If (m_Current Is objItem) Then Exit Sub
        
        m_Current.Draw hWnd, PaintDC, lpRect, pvsNot
        UpdateView 0&, 0&, m_Current
        lpRect.Bottom = 0&
    End If
    
    Set m_Current = Nothing

    If (Not objItem Is Nothing) Then
                
            Set m_Current = objItem
            
        If (m_MenuType = mtcPopup) Then
            m_Current.Draw hWnd, PaintDC, lpRect, pvsSelected
        Else
            m_Current.Draw hWnd, PaintDC, lpRect, pvsHotLight
        End If
        
        UpdateView 0&, 0&, objItem
    End If
    
End Sub

Private Sub ValidateBlank()
    Dim varObj As MenuItem
    
    Dim i As Long
    
    i = m_Col.Count
    
    If (i > 0&) Then
        Set varObj = m_Col(i)
    End If
    
    If varObj Is Nothing Then
        Me.Add "", "$$BLANK$$"
        Exit Sub
    End If
    
    If (varObj.Key = "$$BLANK$$") And (varObj.Caption <> "") Then
        varObj.Key = ""
        
        m_Col.Remove i
        m_Col.Add varObj
        
        Me.Add "", "$$BLANK$$"
    ElseIf (varObj.Key <> "$$BLANK$$") Then
        Me.Add "", "$$BLANK$$"
    End If
    
    If (Item("$$BLANK$$").UserWidth < WrapLimit_Minimum) Then
        Me.Item("$$BLANK$$").UserSize WrapLimit_Minimum
    End If
    
End Sub

Private Sub EditBox_Change()
        
    If (EditItem Is Nothing) Or (Editing = False) Then Exit Sub
        
'    SizeEdit EditItem
    
End Sub

Private Sub EditBox_KeyDown(KeyCode As Long, Shift As Long)
    If (Shift = 0&) Then
        If (KeyCode = vbKeyReturn) Then
            StopEdit False
        ElseIf (KeyCode = vbKeyEscape) Then
            StopEdit True
        End If
    End If
End Sub

Private Sub EditBox_LostFocus()
    If (Editing = True) Then
        StopEdit
        UpdateView
    End If
End Sub

Private Sub EditBox_MouseMove(Button As Long, Shift As Long, x As Long, y As Long)
    Dim iShift As Integer, _
        iButton As Integer
    
    iButton = (Button And &HFFFF&)
    iShift = (Shift And &HFFFF&)
    
    If (Editing = True) Then
        UserControl_MouseMove iButton, iShift, x + EditBox.Left, y + EditBox.Top
    End If

End Sub

Private Sub EditSizer_Resize()
    
    If Not EditItem Is Nothing Then
        EditItem.UserSize EditSizer.Width, EditSizer.Height - 4
        EditSizer.Visible = False
        UserControl_Paint
        SetSizeBox
        EditSizer.Visible = True
    End If

End Sub

Private Sub EditSizer_Rendered()
    
    UpdateView
        
End Sub

Private Sub Scroll_HThumbTrack(ByVal Position As Long)
    
    Scroll.HPosition = Position
    UpdateView 0&, 0&
    
    
End Sub


Private Sub Scroll_VThumbTrack(ByVal Position As Long)

    Scroll.VPosition = Position
    UpdateView 0&, 0&
    
End Sub

Private Sub Scroll_PageDown()

    UpdateView 0&, 60

End Sub

Private Sub Scroll_PageLeft()

    UpdateView -60, 0&

End Sub

Private Sub Scroll_PageRight()

    UpdateView 60, 0&

End Sub

Private Sub Scroll_PageUp()

    UpdateView 0&, -60

End Sub

Private Sub Scroll_LineDown()

    UpdateView 0&, 5

End Sub

Private Sub Scroll_LineLeft()

    UpdateView -5, 0&

End Sub

Private Sub Scroll_LineRight()

    UpdateView 5, 0&

End Sub

Private Sub Scroll_LineUp()

    UpdateView 0&, -5

End Sub

Private Sub m_Font_FontChanged(ByVal PropertyName As String)
    PropertyChanged "Font"
    UserControl_Paint
End Sub

Private Sub ThreadMouse_MouseMove(ByVal x As Long, ByVal y As Long)

    Dim lpRect As RECT

    If PaintNoResize = True Then Exit Sub
    
    GetWindowRect hWnd, lpRect
    
    With lpRect
    
        If (x < .Left) Or (x > .Right) Or _
            (y < .Top) Or (y > .Bottom) Then
            
            If (Not m_Current Is Nothing) Then
                
                lpRect.Bottom = 0&
                
                m_Current.Draw hWnd, PaintDC, lpRect, pvsNot
                UpdateView 0&, 0&, m_Current
            
                Set m_Current = Nothing
            End If
            
        End If
    End With

End Sub

Private Sub UserControl_Paint()
    
    Dim lpRect As RECT, _
        hBrush As Long, _
        lpBrush As LOGBRUSH
    
    Dim Col() As SIZEAPI, _
        StartPos As POINTAPI, _
        lpBands() As Long
    
    Dim tWidth As Long, _
        tHeight As Long
    
    Dim i As Long
    
    Dim varNCLI As NONCLIENTMETRICS
    
    If (PaintNoResize = True) Then Exit Sub
    
    Set m_Current = Nothing
    PaintNoResize = True
    ValidateBlank
    
    i = ConfigureScrollSpace
    
    tWidth = ScaleWidth + m_HScrollSpace
    tHeight = ScaleHeight + m_VScrollSpace
        
    If (PaintDC = 0&) Then
        PaintDC = CreateCompatibleDC(hDC)
    End If
    
    If (PaintBitmap = 0&) Then
        PaintBitmap = CreateCompatibleBitmap(hDC, tWidth, tHeight)
        SelectObject PaintDC, PaintBitmap
    End If
            
    lpRect.Left = 0&
    lpRect.Top = 0&
    lpRect.Right = tWidth
    lpRect.Bottom = tHeight
    lpBrush.lbStyle = BS_SOLID
    lpBrush.lbColor = GetActualColor(m_BackColor)
    
    hBrush = CreateBrushIndirect(lpBrush)
    FillRect PaintDC, lpRect, hBrush
    
    DeleteObject hBrush
            
    If (i > 0&) Then
    
        If (m_MenuType = mtcMenubar) Then
            DrawMenuBar
        Else
            DrawPopupMenu
        End If
        
    End If
    
    If Editing = True Then
        lpRect.Left = 0&
        lpRect.Right = 0&
        
        DrawEditBox lpRect
        EditSizer.Repaint
    End If
    
    UpdateView
        
    PaintNoResize = False
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    Static LastHit As MenuItem
    Dim varObj As MenuItem, _
        htPlace As ItemHitTestPlaceConstants
            
    Dim lpRect As RECT, _
        lpDraw As DRAWITEMSTRUCT
            
    If (Editing = True) Then
        Set varObj = HitTest(x, y, htPlace)
        If varObj Is EditItem Then
        
            varObj.GetLastDraw lpDraw
            varObj.Draw hWnd, PaintDC, lpDraw.rcItem
                        
            If (htPlace = htSubmenu) And (varObj.Submenu.AnyItemVisible = True) Then
                varObj.GetDropRect lpRect
                lpRect.Top = lpDraw.rcItem.Top
                lpRect.Bottom = lpDraw.rcItem.Bottom
                
                Draw3dSq PaintDC, lpRect
            
            ElseIf (htPlace = htPicture) Then
                varObj.GetPictureRect lpRect
                lpRect.Left = lpRect.Left + 1
                lpRect.Right = lpRect.Right - 1
                lpRect.Top = lpRect.Top + 1
                lpRect.Bottom = lpRect.Bottom - 2
                
                Draw3dSq PaintDC, lpRect
            End If
        
            lpRect.Left = 0&
            lpRect.Right = 0&
            
            DrawEditBox lpRect
            
            UpdateView
        Else
            If LastHit Is EditItem Then
            
                LastHit.GetLastDraw lpDraw
                LastHit.Draw hWnd, PaintDC, lpDraw.rcItem
                
                lpRect.Left = 0&
                lpRect.Right = 0&
                
                DrawEditBox lpRect
                
                UpdateView
            End If
        End If
    ElseIf (Sizing = False) Then
        Set varObj = HitTest(x, y)
        SelectItem varObj
    End If
    
    Set LastHit = varObj
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim varObj As MenuItem, _
        lpDraw As DRAWITEMSTRUCT, _
        lpRect As RECT

    On Error Resume Next
    
    If (Button = 1&) And (Shift = 0&) Then
            
        If (Sizing = True) Then
            With EditSizer
            
                If (x < .Left) Or (x > (.Left + .Width)) Or _
                    (y < .Top) Or (y > (.Top + .Height)) Then
                    
                    StopResize
                End If
                
            End With
        Else
            Set varObj = HitTest(x, y)
            
            If (Editing = True) And (Not varObj Is EditItem) Then
                StopEdit
            End If
            
            If varObj Is Nothing Then Exit Sub
                        
            If (m_EditMode = emAuto) Then
                BeginItemEdit varObj
            End If
            
            RaiseEvent ItemSelect(varObj)
        
        End If
    
    ElseIf (Button = 1&) And (GetKeyState(VK_CONTROL) < 0&) Then
    
        If (Editing = False) And (Sizing = False) Then
            If (m_EditMode = emAuto) Then
                Set varObj = HitTest(x, y)
                If varObj Is Nothing Then Exit Sub
                BeginItemResize varObj
            End If
        End If
        
    End If
        
End Sub

Private Sub UserControl_Resize()
    
    If PaintNoResize = True Then Exit Sub
    
    On Error Resume Next
        
    If (UI_Init = False) And (Ambient.UserMode = True) Then
        UI_Init = True
        
        If (ThreadMouse Is Nothing) Then
            Set ThreadMouse = New MouseHook
            ThreadMouse.ThreadHook App.ThreadID
        End If
        
        EditBox.Visible = False
        EditBox.Create hWnd, "", 0&, 0&
    End If
        
    DeleteDC PaintDC
    DeleteObject PaintBitmap
    
    PaintDC = 0&
    PaintBitmap = 0&
    
    UserControl_Paint
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'
    m_MenuType = PropBag.ReadProperty("MenuType", m_MenuType)
    m_BackColor = PropBag.ReadProperty("BackColor", m_BackColor)
    m_RightToLeft = PropBag.ReadProperty("RightToLeft", m_RightToLeft)
    
    Set m_Font = PropBag.ReadProperty("Font", m_Font)
    
    m_Alignment = PropBag.ReadProperty("Alignment", m_Alignment)
    m_Events = PropBag.ReadProperty("Events", m_Events)
    m_EditMode = PropBag.ReadProperty("EditMode", m_EditMode)
    m_DisableNoScroll = PropBag.ReadProperty("DisableNoScroll", m_DisableNoScroll)
    
    m_ScrollTrackPercent = PropBag.ReadProperty("ScrollTrackPercent", m_ScrollTrackPercent)
    
    m_Scrollbars = PropBag.ReadProperty("ScrollBars", m_Scrollbars)
        
    Scroll.Hook hWnd, Ambient.UserMode
    
    Scroll.DisableNoScroll = m_DisableNoScroll
    
    Scroll.Scrollbars = m_Scrollbars
    
    If (m_MenuType = mtcMenubar) Then
        
        Select Case m_Alignment
        
            Case mpaCenter, mpaTop, mpaLeftTop, mpaRightTop, mpaLeft
            
                m_Alignment = mpaTop
                
            Case Else
                
                m_Alignment = mpaBottom
                        
        End Select
        
    End If
    
    UserControl.BackColor = m_BackColor
    UserControl.RightToLeft = m_RightToLeft
    
    UserControl_Paint
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'
    PropBag.WriteProperty "MenuType", m_MenuType
    PropBag.WriteProperty "BackColor", m_BackColor
    PropBag.WriteProperty "RightToLeft", m_RightToLeft
    PropBag.WriteProperty "Font", m_Font

    PropBag.WriteProperty "Events", m_Events
    PropBag.WriteProperty "EditMode", m_EditMode

    PropBag.WriteProperty "Alignment", m_Alignment

    PropBag.WriteProperty "ScrollBars", m_Scrollbars
    PropBag.WriteProperty "DisableNoScroll", m_DisableNoScroll

    PropBag.WriteProperty "ScrollTrackPercent", m_ScrollTrackPercent

End Sub


Private Sub UserControl_Initialize()
'
    Set Scroll = New WndScroll
    
    Set m_Font = New StdFont
    Set m_Col = New Collection
    
    m_MenuType = mtcPopup
    m_Alignment = mpaCenter
    
    m_Font.SIZEAPI = 8
    m_Font.Name = "MS Sans Serif"
    
    m_BackColor = (&H80000000 Or COLOR_DESKTOP)
    
    Set EditBox = New EditWnd
    
    EditBox.UseContainerKeys = True
    EditBox.ThreedBorder = False
    EditBox.SelStart = 0&
    
    EditBox.Style = esMultiline Or esAutoVScroll Or esLeft Or esAutoHScroll
    
    Set EditSizer = New SizeRect
    
    EditSizer.Handles = srAllGrabbers
    EditSizer.GridSize = 4&
    EditSizer.PenStyle = psDot
        
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
        
    Scroll.Unhook
    EditBox.Destroy
    
    Set EditBox = Nothing
    Set ThreadMouse = Nothing
    Set EditSizer = Nothing
    
    Set Scroll = Nothing
    
    If PaintDC <> 0& Then
        DeleteDC PaintDC
    End If
    
    If PaintBitmap <> 0& Then
        DeleteObject PaintBitmap
    End If
    
    Set m_Col = Nothing
    Set m_Font = Nothing
    
End Sub









''' Copyright (C) 2001 Nathan Moschkin

''' ******** NOT FOR PUBLIC COMMERCIAL USE ********
''' Inquire if you would like to use this code commercially.
''' Unauthorized recompilation and/or re-release for commercial
''' use is strictly prohibited.
'''
''' please send changes made to code to me at the address, below,
''' if you plan on making those changes publicly available.

''' e-mail questions or comments to nmosch@tampabay.rr.com







