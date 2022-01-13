VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   2385
   ClientTop       =   2370
   ClientWidth     =   7980
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   7980
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_SysMenu As SystemMenu
Attribute m_SysMenu.VB_VarHelpID = -1

Private WithEvents m_Menu As Menubar
Attribute m_Menu.VB_VarHelpID = -1

Private WithEvents m_WestBar As Sidebar
Attribute m_WestBar.VB_VarHelpID = -1

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Sub Form_Load()
    
    Dim obj As MenuItem, _
        obj2 As MenuItem
    
    Dim i As Long
    
    Set m_SysMenu = New SystemMenu
    Set m_Menu = New Menubar
    
    m_Menu.Create
    
    '' This will switch between fancy, flat (a la Office 2000/XP)
    '' and standard menus with the 3D look (although each item
    '' is completely customizable

    
    Menus.MenuDrawStyle = mdsOfficeXP
    m_SysMenu.Subclass hWnd, True
        
    
   ' Menus.GlobalNoUnicode = True
    
    Set obj = m_Menu.Add("You Should Not See This Item")
    
    obj.Visible = False
    
    Menus.PrefixLib.Add "Ctrl+K"
    
    Set obj2 = obj.Submenu.Add("&Accelerator Test 1", , "TA1")
    obj2.Accelerator.AccelWord = "Ctrl+K, I"
    
    Set obj2 = obj.Submenu.Add("A&ccelerator Test 2", , "TA2")
    obj2.Accelerator.AccelWord = "Ctrl+K, T"
    
    Set obj2 = obj.Submenu.Add("&International: Õמנמרמ Êםידא" + vbCrLf + LoadResString(101), , "TA5")
    
    Set obj2.Visual.Font = New StdFont
    
    obj2.Visual.Font.Name = "Arial Cyr"
    obj2.Visual.Font.Charset = 204&
    
    obj2.SeparatorType = mstCaption
    obj2.Visual.SelectionStyle = mssBevel
    
    obj2.Visual.MultiGradient.Add &HCC0000
    obj2.Visual.MultiGradient.Add &H33AA11
    obj2.Visual.MultiGradient.Add &HAA999
    
    obj2.Visual.ItemForeground = &HDDDDDD
    obj2.Submenu.Add "Example Item"
        
    Set obj2.Submenu.Sidebar.Visual.Font = New StdFont
    
    obj2.Submenu.Sidebar.Visual.Font.Name = "GulimChe"
    obj2.Submenu.Sidebar.Visual.Font.Bold = False
    
    obj2.Submenu.Sidebar.Visual.ItemForeground = vbBlue
    
    obj2.Submenu.Sidebar.Caption = LoadResString(102)
    obj2.Submenu.Sidebar.Escapement = escEastern
    obj2.Submenu.Sidebar.Position = posRight
    
    obj2.Submenu.Sidebar.Visible = True
    
    Set obj2 = obj.Submenu.Add("Accele&rator Test 3", , "TA3")
    obj2.Accelerator.AccelWord = "Ctrl+K, Space"
    
    Set obj2 = obj.Submenu.Add("Accelerator Test &4", , "TA4")
    obj2.Accelerator.AccelWord = "Ctrl+K, N"
    
    Set m_WestBar = obj.Submenu.Sidebar
    
    m_WestBar.Caption = "StdMenu API 4.0"
    m_WestBar.Visible = True
    m_WestBar.Enabled = True
    
    obj.Submenu.Sidebar.Visual.MultiGradient.Add vbInfoBackground
    obj.Submenu.Sidebar.Visual.MultiGradient.Add vbButtonFace
    obj.Submenu.Sidebar.Visual.MultiGradient.Add vbApplicationWorkspace
    
    Set obj.Submenu.Sidebar.Picture = LoadResPicture(103, vbResIcon)
    obj.Submenu.Sidebar.Visual.SizeImage 24, 24
    
    Set obj = m_Menu.Add("&Example Item")
    
    obj.Submenu.Add "Example &1"
    obj.Submenu.Add "Example &2"
    
    obj.Submenu.Add "Exit Command", , "ECAP"
    obj.Submenu.Item("ECAP").Separator = True
    obj.Submenu.Item("ECAP").SeparatorType = mstCaption
           
    obj.Submenu.Item("ECAP").Visual.ItemBkGradient = vbDesktop
    
    Set obj2 = obj.Submenu.Item("ECAP").Submenu.Add("&Exit", , "EXIT")
    obj2.Visual.SelectionStyle = mssHotTrack
                
    m_Menu.SetWindowHandle hWnd
    
    Set obj = m_Menu.Add("&Example Item 2")

    Set obj2 = obj.Submenu.Add("Click To Make Seen What Went Unseen", , "UNSEEN")
    Set obj = obj.Submenu.Add("Caption Goodies")
    obj.SeparatorType = mstCaption
    
    With obj.Visual
        .ItemBackground = &HAA7700
        .ItemBkGradient = &H11AA11
        
        .ItemForeground = vbInfoBackground
    End With
    
    obj.Submenu.Create
    obj.Submenu.MaxHeight = 125
            
    Set obj2 = obj.Submenu.Add("Check Marked Item &A", , "CHECKA")
    obj2.Checkmark = True
    
    Set obj2 = obj.Submenu.Add("Check Marked Item &B", , "CHECKB")
    obj2.Checkmark = True
    obj2.Checked = True
    
    obj.Submenu.Add "-"
    
    Set obj2 = obj.Submenu.Add("Radio Item &1", , "RADIO1")
    obj2.RadioGroup = True
    
    Set obj2 = obj.Submenu.Add("Radio Item &2" + vbCrLf + "&Multi-Lined", , "RADIO2")
    obj2.RadioGroup = True
    obj2.Checked = True
    
    Set obj2 = obj.Submenu.Add("Radio Item &3", , "RADIO3")
    obj2.RadioGroup = True
    
    Set obj = m_Menu.Add("&Help", , "HELP")
    
    obj.Submenu.Create
    
    Set obj2 = obj.Submenu.Add("&About StdMenu 4.0", , "ABOUT")
    Set obj2.Picture = LoadResPicture(106, vbResIcon)
    
    obj.RightAlign = True
    
    Set obj2 = obj.Submenu.Add("&Licensing Information", , "LICENSE")
    Set obj2.Picture = LoadResPicture(105, vbResIcon)
    
    obj2.Visual.SelectionStyle = mssFlat
    
    obj.Submenu.Add "-"
    Set obj2 = obj.Submenu.Add("Right Click This Item", , "WEB")
    Set obj2.Visual.Font = New StdFont
    
    obj2.Visual.Font.Name = "Trebuchet MS"
    obj2.Visual.SelectionStyle = mssHotTrack
    obj2.Visual.DrawBoolPic = True
    
    Set obj2.Picture = LoadResPicture(104, vbResIcon)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set m_Menu = Nothing
    
End Sub

Private Sub m_Menu_ItemRightClick(ByVal Item As StdMenuAPI.MenuItem)

    Select Case Item.Key
        
        Case "WEB"
        
            m_Menu(1).Submenu.Popup hWnd, True
            
        
    End Select

End Sub

Private Sub m_Menu_UserCommand(ByVal Item As MenuItem)
    
    Dim sKey As String, _
        varObj As MenuItem
    
    Dim i As Long
    
    sKey = Item.Key
    
    Select Case sKey
        Case "EXIT"
            Unload Me
            
        Case "UNSEEN"
            Set varObj = m_Menu.Item(1)
            varObj.Visible = (varObj.Visible Xor True)
            
        Case "TA1", "TA2", "TA3", "TA4"
        
            i = Val(Mid(sKey, 3))
            MsgBox "You selected item " & i
            
    End Select
    
End Sub

