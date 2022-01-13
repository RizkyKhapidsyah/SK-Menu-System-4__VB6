VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScaleNumeric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'' ScaleNumeric - Multi-device Scale Mode class (OLE independent)

Option Explicit

'' Conversion Constants

Public Enum UnitModeConstants
    nTwips = vbTwips
    nPoints = vbPoints
    nPixels = vbPixels
    nCharacters = vbCharacters
    nInches = vbInches
    nMillimeters = vbMillimeters
    nCentimeters = vbCentimeters
    nHiMetric = vbHimetric
End Enum


Private Const HIMETRIC_PER_INCH = 2540       ' number HIMETRIC units per inch
Private Const TWIPS_PER_INCH = 1440          ' number TWIP units per inch
Private Const PTS_PER_INCH = 72              ' number points (font size) per inch
Private Const MM_PER_INCH = 25.4             ' number of millimeters per inch
Private Const TWIPS_PER_CHAR_X = 120         ' number TWIP units per logical character unit (Horizontal)
Private Const TWIPS_PER_CHAR_Y = 240         ' number TWIP units per logical character unit (Vertical)

Private m_Value As Double
Private m_Units As UnitModeConstants

Private m_Device As Long

Private m_lpY As Long
Private m_lpX As Long

Public Property Get Value() As Double
    Value = m_Value
End Property

Public Property Let Value(ByVal vData As Double)
    m_Value = vData
End Property

Public Property Get Device() As Long
    Device = m_Device
End Property

Public Property Let Device(ByVal vData As Long)
    
    m_Device = vData
    GetDevicelpXY

End Property

Public Property Get Units() As UnitModeConstants
    
    Units = m_Units

End Property

Public Property Let Units(ByVal vData As UnitModeConstants)
    
    m_Units = vData
    'PropertyChanged "Units"
End Property

Public Property Get AsPixelsX() As Long
    
    Select Case m_Units
        Case nCharacters:
        
            AsPixelsX = CharXToPix(m_Value, m_lpX)
        
        Case Else
        
            AsPixelsX = GetPixValue(m_Value, m_lpX)
    
    End Select
    
End Property

Public Property Let AsPixelsX(ByVal vData As Long)

    Select Case m_Units
    
        Case nCharacters:
            
            m_Value = PixToCharX(vData, m_lpX)
            
        Case Else
        
            m_Value = LetPixValue(vData, m_lpX)
            
    End Select
    
End Property

Public Property Get AsPixelsY() As Long
    
    Select Case m_Units
        Case nCharacters:
        
            AsPixelsY = CharYToPix(m_Value, m_lpY)
        
        Case Else
        
            AsPixelsY = GetPixValue(m_Value, m_lpY)
    
    End Select

End Property

Public Property Let AsPixelsY(ByVal vData As Long)

    Select Case m_Units
    
        Case nCharacters:
            
            m_Value = PixToCharY(vData, m_lpY)
            
        Case Else
        
            m_Value = LetPixValue(vData, m_lpY)
            
    End Select
    
End Property


Public Function DeviceTranslateX(ByVal Value As Double, ByVal FromUnits As UnitModeConstants, ByVal ToUnits As UnitModeConstants) As Double

    Dim v As Double

    If FromUnits = nCharacters Then
        v = CharXToPix(Value, m_lpX)
    Else
        v = GetPixValue(Value, m_lpX, FromUnits)
    End If
    
    If ToUnits = nCharacters Then
        DeviceTranslateX = PixToCharX(v, m_lpX)
    Else
        DeviceTranslateX = LetPixValue(v, m_lpX, ToUnits)
    End If

End Function

Public Function DeviceTranslateY(ByVal Value As Double, ByVal FromUnits As UnitModeConstants, ByVal ToUnits As UnitModeConstants) As Double

    Dim v As Double


    If FromUnits = nCharacters Then
        v = CharYToPix(Value, m_lpY)
    Else
        v = GetPixValue(Value, m_lpY, FromUnits)
    End If
    
    If ToUnits = nCharacters Then
        DeviceTranslateY = PixToCharY(v, m_lpY)
    Else
        DeviceTranslateY = LetPixValue(v, m_lpY, ToUnits)
    End If
    
End Function

Public Function DeviceTranslate(ByVal Value As Double, ByVal FromUnits As UnitModeConstants, ByVal ToUnits As UnitModeConstants) As Double

    Dim v As Double

    If FromUnits = nCharacters Then
        v = (CharXToPix(Value, m_lpX) + CharYToPix(Value, m_lpY) / 2)
    Else
        v = GetPixValue(Value, ((m_lpX + m_lpY) / 2), FromUnits)
    End If
    
    If ToUnits = nCharacters Then
        DeviceTranslate = (PixToCharX(v, m_lpX) + PixToCharY(v, m_lpY)) / 2
    Else
        DeviceTranslate = LetPixValue(v, ((m_lpX + m_lpY) / 2), ToUnits)
    End If

End Function


Private Function PixToHiMetric(ByVal x As Long, ByVal ppli As Long) As Double
    
    PixToHiMetric = MulDiv(HIMETRIC_PER_INCH, x, ppli)

End Function

Private Function HiMetricToPix(ByVal x As Double, ByVal ppli As Long) As Long

    HiMetricToPix = MulDiv(ppli, x, HIMETRIC_PER_INCH)
    
End Function

Private Function LetPixValue(ByVal vData As Double, ByVal lppi As Long, Optional ByVal Unit As UnitModeConstants) As Double
    Dim valReturn As Double, _
        t_Unit As UnitModeConstants
            
    If (Unit <> 0&) Then t_Unit = Unit Else t_Unit = m_Units
        
    Select Case t_Unit
        
        Case nTwips
            valReturn = (vData / lppi) * TWIPS_PER_INCH
        
        Case nInches
            valReturn = vData / lppi
        
        Case nPoints
            valReturn = (vData / lppi) * PTS_PER_INCH
        
        Case nPixels
            valReturn = vData
        
        Case nMillimeters
            valReturn = (vData / lppi) * MM_PER_INCH
        
        Case nCentimeters
            valReturn = (vData / lppi) * (MM_PER_INCH / 10)
        
        Case nHiMetric
            valReturn = (vData / lppi) * HIMETRIC_PER_INCH
    
    End Select
    
    LetPixValue = valReturn

End Function


Private Function GetPixValue(ByVal vData As Double, ByVal lppi As Long, Optional ByVal Unit As UnitModeConstants) As Double
    Dim valReturn As Double, _
        t_Unit As UnitModeConstants
            
    If (Unit <> 0&) Then t_Unit = Unit Else t_Unit = m_Units
        
    Select Case t_Unit
        
        Case nTwips
            valReturn = (vData / TWIPS_PER_INCH) * lppi
        
        Case nInches
            valReturn = vData * lppi
        
        Case nPoints
            valReturn = (vData / PTS_PER_INCH) * lppi
        
        Case nPixels
            valReturn = vData
        
        Case nMillimeters
            valReturn = (vData * lppi) * MM_PER_INCH
        
        Case nCentimeters
            valReturn = (vData * lppi) * (MM_PER_INCH / 10)
        
        Case nHiMetric
            valReturn = (vData / HIMETRIC_PER_INCH) * lppi
        
    End Select
    
    GetPixValue = valReturn
    
End Function

Private Function CharXToPix(ByVal vData As Long, ByVal lppi As Long) As Long
    
    CharXToPix = ((vData * TWIPS_PER_CHAR_X) / TWIPS_PER_INCH) * lppi

End Function

Private Function PixToCharX(ByVal vData As Long, ByVal lppi As Long) As Long
    
    PixToCharX = ((vData / lppi) * TWIPS_PER_INCH) / TWIPS_PER_CHAR_X

End Function

Private Function CharYToPix(ByVal vData As Long, ByVal lppi As Long) As Long

    CharYToPix = ((vData * TWIPS_PER_CHAR_Y) / TWIPS_PER_INCH) * lppi

End Function

Private Function PixToCharY(ByVal vData As Long, ByVal lppi As Long) As Long
    
    PixToCharY = ((vData / lppi) * TWIPS_PER_INCH) / TWIPS_PER_CHAR_Y

End Function

Private Sub GetDevicelpXY()
    Dim hDC As Long
        
    If m_Device = 0& Then
        hDC = GetDC(0&)
        
        m_lpX = GetDeviceCaps(hDC, LOGPIXELSX)
        m_lpY = GetDeviceCaps(hDC, LOGPIXELSY)
        
        ReleaseDC 0&, hDC
        
    Else
        hDC = m_Device
        
        m_lpX = GetDeviceCaps(hDC, LOGPIXELSX)
        m_lpY = GetDeviceCaps(hDC, LOGPIXELSY)
    End If
        
End Sub

Private Sub Class_Initialize()

    GetDevicelpXY
    m_Units = nTwips
    
End Sub

