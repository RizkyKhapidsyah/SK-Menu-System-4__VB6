Attribute VB_Name = "modDraw"
''' GUInerd Standard Menu System
''' Version 4.1

''' Objects/API Dll

''' High-Level Graphics Routines (for menus and menu items)


''' ****************************************************************************
''' *** Check the end of the file for copyright and distribution information ***
''' ****************************************************************************

Option Explicit

Public Enum Rectangles
    rciText = 1&
    rciAccel = 2&
    rciPicture = 3&
    rciArrow = 4&
End Enum

Public Const ResImage_Bullet = 101&
Public Const ResImage_Check = 102&

Public Const ResImage_Close = 103&
Public Const ResImage_Maximize = 104&
Public Const ResImage_Minimize = 105&
Public Const ResImage_Restore = 106&

Public Const ItemPad = 2&
Public Const TextPad = 8&
Public Const ItemAccelSpaces = 12&
Public Const IconicPad = 2&
Public Const IconicSquare = 18&

Public Function MeasureItem(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
    Dim Item As MenuItem, _
        Bar As Sidebar

    Dim Visual As ItemVisualProperties, _
        hDC As Long

    Dim objFind As Object

    Dim lpMeasure As MEASUREITEMSTRUCT, _
        lpSize As SIZEAPI
    
    Dim lpLines() As String, _
        dwAccel As Long
    
    Dim x As Long, _
        y As Long
    
    Dim aText As String
    
    Dim fPic As Boolean, _
        fPicSize As SIZEAPI
    
    Dim fPicComp As SIZEAPI
    
    Dim hFont As Long, _
        oldFont As Long
            
    If (lParam = 0&) Or (hWnd = 0&) Or (IsWindow(hWnd) = 0&) Then Exit Function
        
    CopyMemory lpMeasure, ByVal lParam, Len(lpMeasure)
    
    Set objFind = g_MenuCol.FindMenuItem(lpMeasure.ItemId, hWnd)
    
    If objFind Is Nothing Then Exit Function

    fPicSize.cx = IconicSquare
    fPicSize.cy = IconicSquare

    If TypeOf objFind Is MenuItem Then
        Set Item = objFind
        Set Visual = Item.Visual
    

        If (Item.Checked = True) And (Item.CheckedPicture Is Nothing) Then
            If (Item.RadioGroup = True) Then
                Set Item.CheckedPicture = LoadResPicture(101, vbResBitmap)
            Else
                Set Item.CheckedPicture = LoadResPicture(102, vbResBitmap)
            End If
        End If
    
        If (Not Item.Picture Is Nothing) Or _
            ((Item.Checked = True) And (Not Item.CheckedPicture Is Nothing)) Then
            
            If (Visual.ScaleImages = True) Then
                fPicSize.cx = (IconicPad + Visual.ImageScaleWidth)
                fPicSize.cy = (IconicPad + Visual.ImageScaleHeight)
            Else
                If Not Item.Picture Is Nothing Then
                    fPicSize.cx = ScaleTool.DeviceTranslateX(Item.Picture.Width, nHiMetric, nPixels)
                    fPicSize.cy = ScaleTool.DeviceTranslateY(Item.Picture.Height, nHiMetric, nPixels)
                End If
                
                If Not Item.CheckedPicture Is Nothing Then
                    fPicComp.cx = ScaleTool.DeviceTranslateX(Item.CheckedPicture.Width, nHiMetric, nPixels)
                    fPicComp.cy = ScaleTool.DeviceTranslateY(Item.CheckedPicture.Height, nHiMetric, nPixels)
                End If
                
                If (fPicComp.cx > fPicSize.cx) Then fPicSize.cx = fPicComp.cx
                fPicSize.cx = IconicPad + fPicSize.cx
                
                If (IconicSquare > fPicSize.cx) Then
                    fPicSize.cx = IconicSquare
                End If
                
                If (fPicComp.cy > fPicSize.cy) Then fPicSize.cy = fPicComp.cy
                fPicSize.cy = IconicPad + fPicSize.cy
                
                If (IconicSquare > fPicSize.cy) Then
                    fPicSize.cy = IconicSquare
                End If
                    
            End If
            
            fPic = True
            
        End If
        
        If Not TypeOf Item.Parent Is Menubar Then
            x = fPicSize.cx
            y = fPicSize.cy
            
            '' for the padding on the other side
            x = x + (IconicSquare - IconicPad)
                    
            '' for text spacing
            x = x + (TextPad) + IconicPad
                    
            fPic = True
        End If
        
        hDC = GetDC(hWnd)
        
        hFont = GetItemFont_Decorated(Item, hDC, False)
        oldFont = SelectObject(hDC, hFont)
        
        Item.GetLineInfo lpLines, dwAccel, lpSize, hDC
        
        x = x + lpSize.cx
        If (lpSize.cy > y) Then
            y = lpSize.cy
        End If
                
        If (UBound(lpLines) > 0&) And (y > IconicSquare) Then
            y = y + (ItemPad * 2)
        End If
                
        aText = Item.Accelerator.AccelWord
        
        If (aText <> "") Then
            aText = String(ItemAccelSpaces, " ") + aText
            
            MeasureText_API hDC, aText, lpSize, False
            x = x + lpSize.cx
            If (lpSize.cy > y) Then
                y = lpSize.cy
            End If
            
        End If
                    
        SelectObject hDC, oldFont
        DeleteObject hFont
        
        ReleaseDC hWnd, hDC
        Erase lpLines
        
        If (Item.Separator = True) Then
            If (Item.SeparatorType = mstNormal) Then
                y = (ItemPad * 2&)
                
            ElseIf (Item.SeparatorType = mstCaption) Then
                y = lpSize.cy - 1
                
            End If
        ElseIf Not TypeOf Item.Parent Is Menubar Then
            If (y < IconicSquare) Then y = IconicSquare
            y = y + ItemPad
        End If
        
        If (g_MenuCol.MenuDrawStyle = mdsOfficeXP) Then
            x = x + 2
            y = y + 2
        End If
        
        lpMeasure.itemWidth = x
        lpMeasure.itemHeight = y + 1
                                
    ElseIf TypeOf objFind Is Sidebar Then
        Set Bar = objFind
        Set Visual = Bar.Visual
    
        If (lpMeasure.ItemId = Bar.BreakId) Then
            lpMeasure.itemHeight = 0&
            lpMeasure.itemWidth = 0&
            
            CopyMemory ByVal lParam, lpMeasure, Len(lpMeasure)
            MeasureItem = True
            Exit Function
        End If
    
        If (Not Bar.Picture Is Nothing) Then
            
            If (Visual.ScaleImages = True) Then
                fPicSize.cx = (Visual.ImageScaleWidth)
                fPicSize.cy = (Visual.ImageScaleHeight)
            Else
                If Not Bar.Picture Is Nothing Then
                    fPicSize.cx = ScaleTool.DeviceTranslateX(Bar.Picture.Width, nHiMetric, nPixels)
                    fPicSize.cy = ScaleTool.DeviceTranslateY(Bar.Picture.Height, nHiMetric, nPixels)
                End If
                
                If (IconicSquare > fPicSize.cx) Then
                    fPicSize.cx = IconicSquare
                End If
                
                If (IconicSquare > fPicSize.cy) Then
                    fPicSize.cy = IconicSquare
                End If
                    
            End If
            
            x = fPicSize.cx
            y = fPicSize.cy
                            
            fPic = True
    
        End If
    
        
        hDC = GetDC(hWnd)
        
        hFont = GetItemFont_Decorated(Bar, hDC, False)
        oldFont = SelectObject(hDC, hFont)
        
        Bar.GetLineInfo lpLines, dwAccel, lpSize, hDC
        
        y = y + lpSize.cx + (TextPad * 2)
        
        If (lpSize.cy > x) Then
            x = lpSize.cy
        End If
                
        SelectObject hDC, oldFont
        DeleteObject hFont
        
        ReleaseDC hWnd, hDC
        Erase lpLines
            
        lpMeasure.itemWidth = x
        lpMeasure.itemHeight = y
        
    End If

    Visual.SetMeasure lpMeasure

    CopyMemory ByVal lParam, lpMeasure, Len(lpMeasure)
    MeasureItem = True
    
End Function

Public Function DrawItemIndirect(ByVal Item As Object, ByVal hWnd As Long, _
                                 Optional ByVal SelectState As Long, Optional ByVal fGetDC As Boolean)
                                 
    Dim lpDraw As DRAWITEMSTRUCT, _
        Visual As ItemVisualProperties
        
    Set Visual = Item.Visual
    Visual.GetDraw lpDraw
    
    lpDraw.itemState = lpDraw.itemState And ((ODS_SELECTED + ODS_HOTLIGHT) Xor -1&)
    lpDraw.itemState = lpDraw.itemState + SelectState
    
    lpDraw.hwndItem = hWnd
    
    If (fGetDC = True) Then
        lpDraw.hDC = GetDC(hWnd)
    End If
    
    DrawItem hWnd, ByVal VarPtr(lpDraw)
    
    If (fGetDC = True) Then
        ReleaseDC hWnd, lpDraw.hDC
    End If
    
    Set Visual = Nothing
    
    
End Function

Public Function DrawItem(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
    
    '' local variables
    
    Dim cSep As Boolean, _
        cSel As SelectionStyleConstants
        
    Dim fDisabled As Boolean, _
        fGrayed As Boolean, _
        fMenubar As Boolean, _
        fSelected As Boolean, _
        fPic As Boolean, _
        fRight As Boolean, _
        fXP As Boolean
            
    Dim Item As MenuItem, _
        Bar As Sidebar
    
    Dim Visual As ItemVisualProperties, _
        hDC As Long
    
    Dim objFind As Object

    Dim lpDraw As DRAWITEMSTRUCT, _
        lpSize As SIZEAPI
    
    Dim lpLines() As String, _
        dwAccel As Long
    
    Dim x As Long, _
        y As Long
    
    Dim i As Long, _
        j As Long
    
    Dim a As Long, _
        b As Long
    
    Dim aText As String
    
    Dim cbPicSize As SIZEAPI, _
        cbPicComp As SIZEAPI, _
        cbItem As SIZEAPI
    
    Dim maxLeft As Long, _
        maxRight As Long

    Dim hFont As Long, _
        oldFont As Long, _
        lpFont As LOGFONT, _
        newFont As Long
            
    Dim lpRect As RECT, _
        iPic As StdPicture, _
        fBackColor As Long
    
    Dim bmpDraw As Long, _
        oldBmp As Long
    
    Dim lpBrush As LOGBRUSH, _
        hBrush As Long
    
    Dim lprcPic As RECT
    
    Dim lpPA As POINTAPI, _
        lpPB As POINTAPI

''
On Error Resume Next
''

    
    hDC = g_Handles("_H" + Hex(hWnd)).PaintDC
  
    If (hDC = 0&) Or (lParam = 0&) Or (hWnd = 0&) Or (IsWindow(hWnd) = 0&) Then Exit Function

    CopyMemory lpDraw, ByVal lParam, Len(lpDraw)
    
    With lpDraw.rcItem
        cbItem.cx = (.Right - .Left) + 1&
        cbItem.cy = (.Bottom - .Top) '+ 1&
    End With
    
    bmpDraw = CreateCompatibleBitmap(lpDraw.hDC, cbItem.cx, cbItem.cy)
    oldBmp = SelectObject(hDC, bmpDraw)
    
    If (lpDraw.itemState And ODS_DISABLED) Then
        fDisabled = True
    End If
    
    If (lpDraw.itemState And ODS_GRAYED) Then
        fGrayed = True
    End If
    
    If (lpDraw.itemState And ODS_SELECTED) Then
        fSelected = True
    End If
    
    Set objFind = g_MenuCol.FindMenuItem(lpDraw.ItemId, hWnd)
    
    If objFind Is Nothing Then
        DeleteObject bmpDraw
        SelectObject hDC, oldBmp
        Exit Function
    End If
    
    If TypeOf objFind Is MenuItem Then
        Set Item = objFind
        Set Visual = Item.Visual
        
        If TypeOf Item.Parent Is SystemMenu Then
            Item.GetCtrlState
        End If
                
        Visual.SetDraw lpDraw
        
        If (TypeOf Item.Parent Is Menubar) Then
            cbItem.cx = cbItem.cx - 1
            fMenubar = True
        End If
                
        If ((Item.Separator = True) And (Item.SeparatorType = mstCaption)) Then
            cSep = True
        End If
        
        If ((Visual.SelectionStyle And &HF&) = mssDefault) Then
            If ((cSep = True) Or (fMenubar = True)) Then
                cSel = mssClear
            Else
                cSel = mssColor
            End If
            
            cSel = cSel Or (Visual.SelectionStyle And &HFF0)
        Else
            cSel = Visual.SelectionStyle
        End If
        
        cSel = (cSel Or (Visual.SelectionStyle And mssBevel))
        
        If (Item.RightToLeft = True) Or (Visual.TextAlign = taRight) Then fRight = True
        
        If CBool(lpDraw.itemState And ODS_CHECKED) <> (Item.Checked) Then
            
            Item.Freeze
            Item.Checked = CBool(lpDraw.itemState And ODS_CHECKED)
            Item.Unfreeze
        
        End If
        
        With lpRect
            .Left = 0&
            .Top = 0&
            .Right = cbItem.cx
            .Bottom = cbItem.cy
        End With
        
        lpBrush.lbColor = GetSysColor(COLOR_MENU)
        lpBrush.lbStyle = BS_SOLID
        
        hBrush = CreateBrushIndirect(lpBrush)
        
        FillRect hDC, lpRect, hBrush
        DeleteObject hBrush
                
        If (g_MenuCol.MenuDrawStyle = mdsOfficeXP) And (fMenubar = False) Then
            GetMenuImageMax Item.Parent, maxLeft, maxRight
            cSel = (cSel Or mssFlat)
            fXP = True
            
            If (fSelected = True) Then
                lpBrush.lbColor = &H606060
                lpBrush.lbStyle = BS_SOLID
                
                hBrush = CreateBrushIndirect(lpBrush)
                
                With lpRect
                    .Left = 2&
                    .Right = cbItem.cx
                    .Top = cbItem.cy - 2&
                    .Bottom = cbItem.cy
                End With
                                
                FillRect hDC, lpRect, hBrush
                
                With lpRect
                    .Top = 2&
                    .Bottom = cbItem.cy
                    .Left = cbItem.cx - 2&
                    .Right = cbItem.cx
                End With
                                
                FillRect hDC, lpRect, hBrush
                DeleteObject hBrush
            
            End If
            
            cbItem.cx = cbItem.cx - 2&
            cbItem.cy = cbItem.cy - 2&
        End If
                
        With lpRect
            .Left = 0&
            .Right = cbItem.cx
            
            .Top = 1&
            .Bottom = cbItem.cy - 1&
        End With
        
        lpBrush.lbColor = GetActualColor(Visual.ItemBackground)
        lpBrush.lbStyle = BS_SOLID
        
        hBrush = CreateBrushIndirect(lpBrush)
        
        FillRect hDC, lpRect, hBrush
        DeleteObject hBrush
                
        fBackColor = GetActualColor(Visual.ItemBackground)
                            
        SetBkMode hDC, TRANSPARENT
        
        If (Not Item.Picture Is Nothing) Or _
            ((Item.Checked = True) And (Not Item.CheckedPicture Is Nothing)) Or _
            ((fMenubar = False) And (cSep = False)) Then
            
            If (Not Item.Picture Is Nothing) And (Item.Checked = False) Then
                Set iPic = Item.Picture
            ElseIf (Item.Checked = True) And (Not Item.CheckedPicture Is Nothing) Then
                Set iPic = Item.CheckedPicture
            End If
            
            
            If (Visual.ScaleImages = True) Then
                cbPicSize.cx = (Visual.ImageScaleWidth)
                cbPicSize.cy = (Visual.ImageScaleHeight)
            Else
                If Not iPic Is Nothing Then
                    cbPicSize.cx = ScaleTool.DeviceTranslateX(iPic.Width, nHiMetric, nPixels)
                    cbPicSize.cy = ScaleTool.DeviceTranslateY(iPic.Height, nHiMetric, nPixels)
                Else
                    cbPicSize.cx = IconicSquare - IconicPad
                    cbPicSize.cy = IconicSquare - IconicPad
                End If
            End If
                                
            If (fXP = True) Then
                If (fRight = True) Then
                    cbPicSize.cx = maxRight - ((IconicPad * 2) + ItemPad)
                Else
                    cbPicSize.cx = maxLeft - ((IconicPad * 2) + ItemPad)
                End If
            End If
            
            If (fRight = True) Then
                x = (cbItem.cx - ((IconicPad * 2) + cbPicSize.cx + 2))
            Else
                x = IconicPad
            End If
            
            '' Compute the picture rectangle
            With lprcPic
                .Left = x
                .Right = (.Left + (cbPicSize.cx - 1)) + (IconicPad * 2)
                .Top = 0&
                .Bottom = (.Top + cbItem.cy) - 1&
            End With
            
            If (fXP = True) Then
                
                fBackColor = GetActualColor(Visual.RunnerColor)
                
                lpBrush.lbColor = fBackColor
                lpBrush.lbStyle = BS_SOLID
                
                hBrush = CreateBrushIndirect(lpBrush)
                
                With lpRect
                    .Top = 0&
                    .Bottom = cbItem.cy
                    
                    If (fSelected = False) Then .Bottom = .Bottom + 2
                    
                    If (maxLeft > 0&) Then
                        .Left = 2&
                        .Right = maxLeft
                    
                        FillRect hDC, lpRect, hBrush
                    End If
                
                    If (maxRight > 0&) Then
                        .Right = cbItem.cx '+ 2
                        .Left = .Right - maxRight
                        
                        FillRect hDC, lpRect, hBrush
                    End If
                End With
                
                DeleteObject hBrush
                
                
                If (fSelected = True) Then
                    fBackColor = GetActualColor(Visual.RunnerSelection)
                End If
            End If
                            
            If Not iPic Is Nothing Then
                
                '' Compute the fill area for the picture rectangle
                                                                                        
                CopyMemory lpRect, lprcPic, 16&
                
                With lpRect
                    .Left = .Left + 1
                    .Top = .Top + 1
                    
                    '' don't do that to right and bottom because
                    '' of DC compat mode (FillRect() does not fill the last
                    '' x and y of a rectangle on most screen DCs)
                End With
                    
                If (Item.Checked = True) And (fSelected = True) And (fDisabled = False) And _
                    ((cSel And mssNoCheckBevel) = 0&) Then
                    
                    fBackColor = GetActualColor(Visual.ItemCheckBackground)
                End If
                
                lpBrush.lbColor = fBackColor
                lpBrush.lbStyle = BS_SOLID
                
                hBrush = CreateBrushIndirect(lpBrush)
                
                FillRect hDC, lpRect, hBrush
                DeleteObject hBrush
                
                '' computer the picture extent
                
                With lpRect
                    .Left = .Left - 1
                    .Top = .Top - 1
                    
                    x = (.Right - .Left) + 1&
                    .Left = .Left + ((x / 2&) - (cbPicSize.cx / 2&))
                    
                    x = (.Bottom - .Top) + 1&
                    .Top = .Top + ((x / 2&) - (cbPicSize.cy / 2&))
                    
                    .Right = .Left + (cbPicSize.cx - 1)
                    .Bottom = .Top + (cbPicSize.cy - 1)
                End With
                                                        
                If (fSelected = True) Or (Visual.DrawBoolPic = False) Then
                    
                    If (fSelected = True) And (fXP = True) Then
                        Render_Picture hDC, iPic, lpRect, fGrayed, fBackColor, , True
                    Else
                        Render_Picture hDC, iPic, lpRect, fGrayed, fBackColor
                    End If
                Else
                    Render_Picture hDC, iPic, lpRect, fGrayed, fBackColor, True
                End If
                                
                CopyMemory lpRect, lprcPic, 16&
                
                With lpRect
                    .Bottom = .Bottom - 1&
                End With
                
                If (cSel And mssFlat) Then
                
                    If (fSelected = True) Then
                
                        FlatBox hDC, lpRect
                    End If
                    
                ElseIf (fDisabled = False) And _
                   (fSelected = True) And ((Item.Checked = False) _
                   Or ((cSel And mssNoCheckBevel) <> 0&)) Then
                        
                    lpRect.Top = lpRect.Top + 1
                    Draw3DBox_Raised hDC, lpRect
                
                ElseIf (Item.Checked = True) And ((cSel And mssNoCheckBevel) = 0&) Then
                    
                    Draw3DBox_Sunken hDC, lpRect
                    
                End If
                
            End If
                
            fPic = True
        End If
            
        If (fPic = True) Then
            With lpRect
                .Top = 1
                .Bottom = cbItem.cy - 1
                
                If (iPic Is Nothing) Or (fXP = True) Then
                    .Left = maxLeft
                    .Right = cbItem.cx - maxRight
                Else
                    If (Item.RightToLeft = True) Or (Visual.TextAlign = taRight) Then
                        .Left = ItemPad
                        .Right = (cbItem.cx - 1) - (cbPicSize.cx + (IconicPad * 2) + (TextPad / 4))
                    Else
                        .Left = (cbPicSize.cx + (IconicPad * 2) + (TextPad / 4))
                        .Right = cbItem.cx
                    End If
                End If
                
            End With
                            
        End If
                
        If (fGrayed = True) Then
            lpBrush.lbColor = GetSysColor(COLOR_MENU)
            lpBrush.lbStyle = BS_SOLID
            
            hBrush = CreateBrushIndirect(lpBrush)
            FillRect hDC, lpRect, hBrush
            
            DeleteObject hBrush
            
        Else
            
            If (((cSel And &HF&) = mssColor) And (fSelected = True)) Then
                                
                If (Visual.SelMultiGradient.Count > 0&) Then
                    
                    MultiGradFill hDC, Visual.SelMultiGradient, lpRect
                
                ElseIf (Visual.SelectBkGradient <> -1&) Then
                
                    GradientFill hDC, GetActualColor(Visual.SelectBackground), _
                                      GetActualColor(Visual.SelectBkGradient), _
                                      lpRect, False
                                
                Else
                    lpBrush.lbColor = GetActualColor(Visual.SelectBackground)
                    lpBrush.lbStyle = BS_SOLID
                    
                    hBrush = CreateBrushIndirect(lpBrush)
                    FillRect hDC, lpRect, hBrush
                    
                    DeleteObject hBrush
                End If
                
                SetTextColor hDC, GetActualColor(Visual.SelectForeground)
            Else
                
                If (Visual.MultiGradient.Count > 0&) Then
                    MultiGradFill hDC, Visual.MultiGradient, lpRect
                
                ElseIf (Visual.ItemBkGradient <> -1&) Then
                    
                    GradientFill hDC, GetActualColor(Visual.ItemBackground), _
                                      GetActualColor(Visual.ItemBkGradient), _
                                      lpRect, False
                                
                Else
                    lpBrush.lbColor = GetActualColor(Visual.ItemBackground)
                    lpBrush.lbStyle = BS_SOLID
                    
                    hBrush = CreateBrushIndirect(lpBrush)
                    FillRect hDC, lpRect, hBrush
                    
                    DeleteObject hBrush
                End If
            
                If ((fSelected = False) Or (cSel = mssClear)) Then
                    SetTextColor hDC, GetActualColor(Visual.ItemForeground)
                ElseIf (fSelected = True) Then
                    SetTextColor hDC, GetActualColor(Visual.SelectForeground)
                End If
                
            End If
        End If
                
        With lpRect
        
            If Not iPic Is Nothing Then
                .Top = lprcPic.Top
                .Bottom = lprcPic.Bottom - 1&
            Else
                .Top = 1&
                .Bottom = cbItem.cy - 1&
            End If
            
            If (fRight = True) Then
                If Not iPic Is Nothing Then
                    .Right = lprcPic.Right
                Else
                    .Right = cbItem.cx - ItemPad
                End If
            
                .Left = 1&
            Else
                If Not iPic Is Nothing Then
                    .Left = lprcPic.Left
                Else
                    .Left = 1&
                End If
            
                .Right = cbItem.cx - ItemPad
            End If
            
            If (fMenubar = True) And (iPic Is Nothing) Then
                .Top = 0&
                .Right = .Right + 1
            End If
                
        End With
        
        If ((cSel And &HF&) = mssClear) Or ((cSel And mssFlat) = mssFlat) Then
            If (lpDraw.itemState And ODS_HOTLIGHT) Or _
                ((cSep = True) And (fSelected = False)) Then
                            
                With lpRect
                    .Left = 0&
                    .Top = 0&
                End With
                            
                If (cSel And mssFlat) = mssFlat Then
                    FlatBox hDC, lpRect
                ElseIf (cSel And mssBevel) Then
                    BevelBox hDC, lpRect
                Else
                    Draw3DBox_Raised hDC, lpRect
                End If
            
            ElseIf (fSelected = True) And ((cSep = False) Or (Item.Submenu.AnyItemVisible = True)) Then
                
                With lpRect
                    .Left = 0&
                    .Top = 0&
                End With
                            
                If (cSel And mssFlat) = mssFlat Then
                    FlatBox hDC, lpRect
                    
                ElseIf (cSel And mssBevel) Then
                    BevelBox hDC, lpRect, True
                Else
                    Draw3DBox_Sunken hDC, lpRect
                End If
                
            End If
        
        End If
            
        If (Item.Caption = "-") Then
            
            lpPA.y = (cbItem.cy / 2&)
            lpPB.y = lpPA.y
            
            lpPA.x = lpRect.Left
            lpPB.x = lpRect.Right
                                    
            If Not iPic Is Nothing Then
                If fRight = True Then
                    lpPB.x = lpPB.x - maxRight
                Else
                    lpPA.x = lpPA.x + maxLeft
                End If
            End If
            
            BevelLine hDC, lpPA, lpPB, False, False
            
        End If
            
        hFont = GetItemFont_Decorated(Item, hDC, ((fSelected = True) And (fGrayed = False)))
        oldFont = SelectObject(hDC, hFont)
        
        Item.GetLineInfo lpLines, dwAccel, lpSize, hDC
        SetTextAlign hDC, TA_LEFT
                        
        i = -1&
        i = UBound(lpLines)
        
        If (i <> -1&) Then
            
            If (cbItem.cy > lpSize.cy) Then
                y = ((cbItem.cy / 2&) - (lpSize.cy / 2&)) - (ItemPad / 2)
            Else
                y = 0&
            End If
                                     
            If (fMenubar = True) Then
                y = y + 1
            End If
            
            For j = 0 To i
            
                MeasureText_API hDC, lpLines(j), lpSize, True
                
                Select Case Visual.TextAlign
                    Case taLeft
                        If (fMenubar = True) And (fPic = False) Then
                            x = ItemPad * 2&
                        Else
                            x = (IconicPad / 2) + TextPad + (cbPicSize.cx)
                        End If
                    
                    Case taCenter
                        x = (cbItem.cx / 2&) - (lpSize.cx / 2&)
                    
                    Case taRight
                        If (fMenubar = True) And (fPic = False) Then
                            x = cbItem.cx - ((ItemPad * 2&) + lpSize.cx)
                        Else
                            x = cbItem.cx - ((IconicPad / 2) + TextPad + lpSize.cx + cbPicSize.cx)
                        End If
                    
                End Select
                                    
                a = InStr(1, lpLines(j), "&")
                
                If (a <> 0&) And ((lpDraw.itemState And ODS_NOACCEL) = 0&) Then
                    If (a <> 1) Then
                        aText = Mid(lpLines(j), 1, a - 1)
                        Draw_And_Advance aText, hDC, x, y, ((fGrayed Or fDisabled) Xor True), taLeft, False
                    End If
                    
                    GetObject hFont, Len(lpFont), lpFont
                    lpFont.lfUnderline = 1&
                    
                    newFont = CreateFontIndirect(lpFont)
                    SelectObject hDC, newFont
                    
                    aText = Mid(lpLines(j), a + 1, 1)
                    
                    If (a + 2) <= Len(lpLines(j)) Then
                        Draw_And_Advance aText, hDC, x, y, ((fGrayed Or fDisabled) Xor True), taLeft, False
                    Else
                        Draw_And_Advance aText, hDC, x, y, ((fGrayed Or fDisabled) Xor True), taLeft, True
                    End If
                    
                    SelectObject hDC, hFont
                    DeleteObject newFont
                    
                    If (a + 2) <= Len(lpLines(j)) Then
                        aText = Mid(lpLines(j), a + 2)
                        Draw_And_Advance aText, hDC, x, y, ((fGrayed Or fDisabled) Xor True), taLeft, True
                    End If
                    
                Else
                    aText = GetPlainCaption(lpLines(j))
                    Draw_And_Advance aText, hDC, x, y, ((fGrayed Or fDisabled) Xor True), taLeft, True
                End If
            
            Next j
        
        End If
        
        aText = Item.Accelerator.AccelWord
    
        If (aText <> "") And (Visual.TextAlign <> taCenter) Then
            MeasureText_API hDC, aText, lpSize, False
                        
            If (Visual.TextAlign = taLeft) Then
                x = (cbItem.cx - (IconicSquare + (ItemPad / 2))) - lpSize.cx
            Else
                x = (ItemPad / 2) + IconicSquare
            End If
        
            y = ((cbItem.cy / 2&) - (lpSize.cy / 2&)) - (ItemPad / 2)
            Draw_And_Advance aText, hDC, x, y, ((fGrayed Or fDisabled) Xor True), taLeft, True, False
        
        End If
    
        SelectObject hDC, oldFont
        DeleteObject hFont
        
        Set Item = Nothing
        Set iPic = Nothing
        
    ElseIf TypeOf objFind Is Sidebar Then
        Set Bar = objFind
        Set Visual = Bar.Visual
    
        If ((Visual.SelectionStyle And &HF&) = mssDefault) Then
            cSel = mssHotTrack
        Else
            cSel = Visual.SelectionStyle
        End If
        
        If (lpDraw.ItemId = Bar.BreakId) Then
            DrawItem = True
            Set Bar = Nothing
            
            Exit Function
            
        End If
    
        If (fDisabled = True) Then fSelected = False
        If (Bar.RightToLeft = True) Or (Visual.TextAlign = taRight) Then fRight = True
        
        Visual.SetDraw lpDraw
        
        lpRect.Right = cbItem.cx
        lpRect.Bottom = cbItem.cy
        
        lpBrush.lbColor = GetActualColor(Visual.ItemBackground)
        lpBrush.lbStyle = BS_SOLID
        
        hBrush = CreateBrushIndirect(lpBrush)
        
        FillRect hDC, lpRect, hBrush
        DeleteObject hBrush
        
        SetBkMode hDC, TRANSPARENT
        
        hFont = GetItemFont_Decorated(Bar, hDC, fSelected)
        oldFont = SelectObject(hDC, hFont)
        
        If Not Bar.Picture Is Nothing Then
            Set iPic = Bar.Picture
        End If
        
        If (Visual.ScaleImages = True) Then
            cbPicSize.cx = (Visual.ImageScaleWidth)
            cbPicSize.cy = (Visual.ImageScaleHeight)
        Else
            If Not iPic Is Nothing Then
                
                cbPicSize.cx = ScaleTool.DeviceTranslateX(iPic.Width, nHiMetric, nPixels)
                cbPicSize.cy = ScaleTool.DeviceTranslateY(iPic.Height, nHiMetric, nPixels)
            
            Else
                cbPicSize.cx = IconicSquare - IconicPad
                cbPicSize.cy = IconicSquare - IconicPad
                
            End If
            
        End If
            
        lpRect.Left = 0&
        lpRect.Top = 0&
        lpRect.Right = cbItem.cx
        lpRect.Bottom = cbItem.cy
                    
        If (fSelected = True) Then
            a = Visual.MultiGradient.Count
            If (a <> 0&) Then
                fBackColor = GetActualColor(Visual.MultiGradient.Color(a))
            
            ElseIf (Visual.SelectBkGradient = -1&) Then
                fBackColor = GetActualColor(Visual.SelectBackground)
            Else
                fBackColor = GetActualColor(Visual.SelectBkGradient)
            End If
        Else
            If (Visual.ItemBkGradient = -1&) Then
                fBackColor = GetActualColor(Visual.ItemBackground)
            Else
                fBackColor = GetActualColor(Visual.ItemBkGradient)
            End If
        End If
                            
        If (fSelected = True) And _
            ((lpDraw.itemState And ODS_GRAYED) = 0&) And _
            ((Visual.SelectionStyle <> mssClear) Or (Visual.MultiGradient.Count <> 0&)) Then
                                                        
            If (Visual.MultiGradient.Count > 0&) Then
                MultiGradFill hDC, Visual.MultiGradient, lpRect, True
            
            ElseIf (Visual.SelectBkGradient <> -1&) Then
                GradientFill hDC, _
                    Visual.SelectBackground, Visual.SelectBkGradient, lpRect, True
            Else
                lpBrush.lbColor = GetActualColor(Visual.SelectBackground)
                lpBrush.lbStyle = BS_SOLID
                
                hBrush = CreateBrushIndirect(lpBrush)
                
                FillRect hDC, lpRect, hBrush
                DeleteObject hBrush
            End If
                                
            SetTextColor hDC, GetActualColor(Visual.SelectForeground)
            
        Else
            
            If (Visual.ItemBkGradient <> -1&) Then
                GradientFill hDC, _
                    Visual.ItemBackground, Visual.ItemBkGradient, lpRect, True
            End If
            
            SetTextColor hDC, GetActualColor(Visual.ItemForeground)
            SetBkColor hDC, fBackColor
            
        End If
        
        If Not iPic Is Nothing Then
                
            y = cbItem.cy - (IconicPad + ItemPad + cbPicSize.cy)
            x = ((cbItem.cx / 2&) - ((cbPicSize.cx + IconicPad) / 2&))
            
            With lpRect
                .Left = x + 1
                .Right = x + cbPicSize.cx + 2
                
                .Top = y
                .Bottom = y + cbPicSize.cy + 2
            End With
                            
            With lpRect
                .Left = .Left + 1
                .Right = .Right - 1
                .Bottom = .Bottom - 2
                .Top = .Top + 1
            End With
                                                    
            Render_Picture hDC, iPic, lpRect, False, fBackColor
            
            With lpRect
                .Left = .Left - (IconicPad / 2)
                .Top = .Top - (IconicPad / 2)
                .Right = .Right + (ItemPad * 2)
                .Bottom = .Bottom + (ItemPad * 2)
            End With
            
            
            cbPicSize.cx = cbPicSize.cx + IconicPad
            cbPicSize.cy = cbPicSize.cy + IconicPad
            
            fPic = True
        End If
        
        If (fPic = True) Then
            With lpRect
                .Left = 0
                .Right = cbItem.cx
                .Top = (ItemPad / 2)
                
                If iPic Is Nothing Then
                    .Bottom = cbItem.cy - (ItemPad / 2)
                Else
                    .Bottom = cbItem.cy - (cbPicSize.cy)
                End If
                
            End With
        End If
                    
        If ((fSelected = True) Or (lpDraw.itemState And ODS_HOTLIGHT)) And _
            ((cSel = mssClear) Or (cSel = mssFlat)) Then
        
            lpRect.Right = lpRect.Right - 1
            lpRect.Bottom = lpRect.Bottom - 1
        
            If (lpDraw.itemState And ODS_HOTLIGHT) Then
                Draw3DBox_Raised hDC, lpRect
            Else
                Draw3DBox_Sunken hDC, lpRect
            End If
        
            lpRect.Right = lpRect.Right + 1
            lpRect.Bottom = lpRect.Bottom + 1
            
            If (cSel = mssFlat) Then
                FlatBox hDC, lpRect
            End If
    
        End If
        
        Bar.GetLineInfo lpLines, dwAccel, lpSize, hDC
    
        If (Bar.Escapement = esc90) Then
            x = ((cbItem.cx / 2) - (lpSize.cy / 2))
            y = lpRect.Bottom - (TextPad + ItemPad)
        ElseIf (Bar.Escapement = esc270) Then
            x = cbItem.cx - ((cbItem.cx / 2) - (lpSize.cy / 2))
            y = lpRect.Top + (TextPad + ItemPad)
        End If
    
        i = -1&
        i = UBound(lpLines)
        
        a = (lpRect.Bottom - lpRect.Top) + 1
        
        For j = 0 To i
            aText = GetPlainCaption(lpLines(j))
            MeasureText_API hDC, aText, lpSize
            
            If (Bar.Escapement = esc90) Then
                Select Case Visual.TextAlign
                    Case taLeft:
                        y = lpRect.Bottom - (TextPad + ItemPad)
                    
                    Case taCenter:
                        y = lpRect.Bottom - ((a / 2) - (lpSize.cx / 2))
                        
                    Case taRight:
                        y = lpRect.Top + ((TextPad + ItemPad) + lpSize.cx)
                    
                End Select
            
            ElseIf (Bar.Escapement = esc270) Then
                Select Case Visual.TextAlign
                    Case taRight:
                        y = lpRect.Top + (TextPad + ItemPad)
                        
                    Case taCenter:
                        y = lpRect.Top + ((a / 2) - (lpSize.cx / 2))
                                        
                    Case taLeft:
                        y = lpRect.Bottom - ((TextPad + ItemPad) + lpSize.cx)
                    
                    
                End Select
                
            End If
            
            TextOut_API hDC, aText, x, y
            
            If (Bar.Escapement = esc90) Then
                x = x + lpSize.cy
            Else
                x = x - lpSize.cy
            End If
        Next j
    
        SelectObject hDC, oldFont
        DeleteObject hFont
        
    End If

    If (fXP = True) Then
        cbItem.cx = cbItem.cx + 2
        cbItem.cy = cbItem.cy + 2
    End If
    
    x = lpDraw.rcItem.Left
    y = lpDraw.rcItem.Top
    
    BitBlt lpDraw.hDC, x, y, cbItem.cx, cbItem.cy, hDC, 0&, 0&, SRCCOPY

    SelectObject hDC, oldBmp
    DeleteObject bmpDraw
        
    Set Bar = Nothing
    Set Item = Nothing
        
    Set Visual = Nothing
    Erase lpLines
        
    DrawItem = True
    
End Function

Public Function Draw_And_Advance(ByVal lpszText As String, ByVal hDC As Long, x As Long, y As Long, ByVal Enabled As Boolean, ByVal TextAlign As TextAlignConstants, Optional ByVal AdvanceLine As Boolean, Optional ByVal Escaped As Boolean)
    Dim lpSize As SIZEAPI

    MeasureText_API hDC, lpszText, lpSize
    
    If (Enabled = False) Then
        DrawDisabled_API hDC, lpszText, x, y
    Else
        TextOut_API hDC, lpszText, x, y
    End If
    
    If Escaped = False Then
    
        If (AdvanceLine = True) Then
            y = y + lpSize.cy
        Else
            x = x + lpSize.cx
        End If
    Else
        If (AdvanceLine = True) Then
            x = x + lpSize.cy
        Else
            y = y + lpSize.cx
        End If
    End If
    
End Function

Public Function GetItemFont_Decorated(ByVal Item As Object, ByVal hDC As Long, Optional ByVal Selected As Boolean) As Long
    Dim lpFont As LOGFONT, _
        hFont As Long
                    
    ' Get information about the current font handle selected
    ' for the device context of the menu frame.
        
    ' get the object LOGFONT structure from the font handle
            
    hFont = GetItemFont(hDC, Item)
    GetObject hFont, Len(lpFont), lpFont
        
    If TypeOf Item Is MenuItem Then
        ' change the weight to bold for default items.
                    
        If (Item.Default = True) Then
            lpFont.lfWeight = 700&
        End If
        
        ' 600 = DEMi/SEMi Bold : For caption Separators.
        
        If (Item.SeparatorType And mstCaption) Then
            lpFont.lfWeight = 600
            lpFont.lfHeight = lpFont.lfHeight + 1.5
            lpFont.lfUnderline = 0&
        End If
    End If
                    
    If (Item.Visual.SelectionStyle = mssHotTrack) And (Selected = True) Then
        lpFont.lfUnderline = True
    End If
    
    If TypeOf Item Is Sidebar Then
        Select Case Item.Escapement
        
            Case esc90:
            
                lpFont.lfEscapement = 900&
                lpFont.lfOrientation = 900&
            
            Case esc270:
            
                lpFont.lfEscapement = 2700&
                lpFont.lfOrientation = 2700&
            
        End Select
        
        lpFont.lfQuality = 2&
        lpFont.lfHeight = (lpFont.lfHeight)
    End If
    
    ' Create a new font handle from the modified LOGFONT
    ' and set it active on the drawing device.
    
    DeleteObject hFont
    
    GetItemFont_Decorated = CreateFontIndirect(lpFont)

End Function

Public Function GetItemFont(hDC As Long, ByVal Item As Object) As Long
    On Error Resume Next
    
    Dim objFont As StdFont, _
        lfData As LOGFONT, _
        hFont As Long
        
    Dim varItem As MenuItem, _
        varSidebar As Sidebar
    
    If TypeOf Item Is MenuItem Then
        Set varItem = Item
        
        Set objFont = varItem.Visual.Font
        
        If (objFont Is Nothing) And (varItem.Parent.Font Is Nothing) Then
            Set objFont = g_SysMenuFont
        ElseIf (objFont Is Nothing) Then
            Set objFont = varItem.Parent.Font
        End If
        
        Set varItem = Nothing
        
    ElseIf TypeOf Item Is Sidebar Then
        Set varSidebar = Item
        
        Set objFont = varSidebar.Visual.Font
        
        If (objFont Is Nothing) And (varSidebar.Parent.Font Is Nothing) Then
            Set objFont = g_SysMenuFont
        ElseIf (objFont Is Nothing) Then
            Set objFont = varSidebar.Parent.Font
        End If
        
        Set varSidebar = Nothing
    End If
            
    GetLogFont objFont, lfData, hDC
    
    If Item.Default = True Then
        lfData.lfWeight = 700&
    End If
    
    GetItemFont = CreateFontIndirect(lfData)
    
End Function

Public Function GetMenuImageMax(Menu As Object, cbLeft As Long, cbRight As Long, Optional lpfXP As Boolean) As Long
    
    cbLeft = Menu.ImageMax_Left
    cbRight = Menu.ImageMax_Right
    
    If (cbLeft <> 0&) Then
        cbLeft = cbLeft + (IconicPad * 2) + ItemPad
    End If
    
    If (cbRight <> 0&) Then
        cbRight = cbRight + (IconicPad * 2) + ItemPad
    End If
    
    If IsMissing(lpfXP) = False Then
        If (g_MenuCol.MenuDrawStyle = mdsOfficeXP) Then
            lpfXP = True
        Else
            lpfXP = False
        End If
    End If

End Function

Public Function GetSubmenuCursor(ByVal hDC As Long, Optional ByVal cx As Long = 8&, Optional ByVal cy As Long = 8&, Optional ByVal fCursorLeft As Boolean, Optional ByVal crColor As Long) As Long
    Dim hBitmap As Long, _
        hDrawDC As Long
        
    Dim hPen As Long, _
        lpRect As RECT, _
        hBrush As Long, _
        lpBrush As LOGBRUSH
        
    Dim hOldPen As Long, _
        lpPoint As POINTAPI
       
    hDrawDC = CreateCompatibleDC(hDC)
    
    hBitmap = CreateCompatibleBitmap(hDC, cx, cy)
    SelectObject hDrawDC, hBitmap
    
    lpBrush.lbStyle = BS_SOLID
    lpBrush.lbColor = &HC0C0C0
    
    hBrush = CreateBrushIndirect(lpBrush)
    
    lpRect.Right = cx + 1
    lpRect.Bottom = cy + 1
    
    FillRect hDrawDC, lpRect, hBrush
    
    DeleteObject hBrush
    
    hPen = CreatePen(PS_SOLID, 1&, crColor)
    hOldPen = SelectObject(hDrawDC, hPen)
    
    If (fCursorLeft = True) Then
    
        MoveToEx hDrawDC, cx - 2, 2, lpPoint
        LineTo hDrawDC, cx - 2, cy
        LineTo hDrawDC, (cx / 4), ((cy - 1) / 2&)
        LineTo hDrawDC, cx - 2, 2
            
    Else
    
        MoveToEx hDrawDC, 2, 2, lpPoint
        LineTo hDrawDC, 2, cy
        LineTo hDrawDC, cx - (cx / 4), ((cy - 1) / 2&)
        LineTo hDrawDC, 2, 2
    
    End If
                
    SelectObject hDrawDC, hOldPen
    DeleteObject hPen
    
    lpBrush.lbColor = crColor
    hBrush = CreateBrushIndirect(lpBrush)
    SelectObject hDrawDC, hBrush
        
    FloodFill hDrawDC, (cx / 2), (cy / 2), crColor
    
    DeleteObject hBrush
    DeleteDC hDrawDC
    
    GetSubmenuCursor = hBitmap
    
End Function







''' Copyright (C) 2001 Nathan Moschkin

''' ****************** NOT FOR COMMERCIAL USE *****************
''' Inquire if you would like to use this code commercially.
''' Unauthorized recompilation and/or re-release for commercial
''' use is strictly prohibited.
'''
''' please send changes made to code to me at the address, below,
''' if you plan on making those changes publicly available.

''' e-mail questions or comments to nmosch@tampabay.rr.com





