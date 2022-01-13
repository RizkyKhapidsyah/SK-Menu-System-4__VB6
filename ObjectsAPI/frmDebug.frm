VERSION 5.00
Begin VB.Form frmDebug 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Debug Window"
   ClientHeight    =   2400
   ClientLeft      =   270
   ClientTop       =   9630
   ClientWidth     =   6585
   Icon            =   "frmDebug.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDbg 
      Height          =   525
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   990
      Width           =   1245
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_SysMenu As SystemMenu
Attribute m_SysMenu.VB_VarHelpID = -1


Private Sub Form_Load()

    Set m_SysMenu = New SystemMenu
    m_SysMenu.Subclass hWnd

End Sub

Private Sub Form_Resize()
    txtDbg.Move 0, 0, ScaleWidth, ScaleHeight
End Sub
