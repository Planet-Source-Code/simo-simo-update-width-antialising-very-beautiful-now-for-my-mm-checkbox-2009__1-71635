VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LineGS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"SmoothLineDIB"
Option Explicit

'Original TMT Pascal/Asm code by Jonas Widarsson
'
'Implemented in Vb6 by Dana Seaman
'Send comments/bug reports to dseaman@ieg.com.br
'
Public Enum cThickness
   Thin
   Thick
End Enum
Private Type RGBQUAD
   Blue                 As Byte
   Green                As Byte
   Red                  As Byte
   Reserved             As Byte
End Type

Private Type BITMAPINFOHEADER
   biSize               As Long
   biWidth              As Long
   biHeight             As Long
   biPlanes             As Integer
   biBitCount           As Integer
   biCompression        As Long
   biSizeImage          As Long
   biXPelsPerMeter      As Long
   biYPelsPerMeter      As Long
   biClrUsed            As Long
   biClrImportant       As Long
End Type

Private Type BITMAPINFO
   bmiHeader            As BITMAPINFOHEADER
End Type

Private Type RECT
   Left     As Long
   Top      As Long
   Right    As Long
   Bottom   As Long
End Type

Private Const DIB_RGB_COLORS As Long = 0
Private Const Pi        As Single = 3.141592
Private Const HalfPi    As Single = Pi / 2
Private Const cThin     As Single = Pi * 0.34
Private Const cThick    As Single = Pi * 0.17
Private Const Rads      As Single = Pi / 180
Private Const PS_SOLID  As Long = 0

Private Binfo           As BITMAPINFO
Private buf()           As RGBQUAD
Private InDIBits        As Boolean
Private Red             As Long
Private Green           As Long
Private Blue            As Long
Private m_Color         As Long
Private m_hDC           As Long
Private m_W1            As Long
Private m_H1            As Long
Private m_Handle        As Long

Private Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

'Public Sub Widget(rct As RECT)

'End Sub
Private Function TranslateColour(ByVal clr As OLE_COLOR, _
   Optional hPal As Long = 0) As Long
   If OleTranslateColor(clr, hPal, TranslateColour) Then
      TranslateColour = vbBlack 'CLR_INVALID
   End If
End Function

Public Sub DIB(ByVal hdc As Long, ByVal Handle As Long, ByVal W1 As Long, ByVal H1 As Long)
   m_hDC = hdc
   m_Handle = Handle
   m_W1 = W1
   m_H1 = H1
   Pic2Array
End Sub

Private Sub Pic2Array()
   ReDim buf(0 To (m_W1 - 1), m_H1 - 1) As RGBQUAD
   With Binfo.bmiHeader
      .biSize = 40
      .biWidth = m_W1
      .biHeight = -m_H1
      .biPlanes = 1
      .biBitCount = 32
      .biCompression = 0
      .biClrUsed = 0
      .biClrImportant = 0
      .biSizeImage = m_W1 * m_H1
   End With
   'Copy hDC to Array
   GetDIBits m_hDC, m_Handle, 0, m_H1, buf(0, 0), Binfo, DIB_RGB_COLORS
   'Set local flag
   InDIBits = True
End Sub

Public Sub CircleGP(ByVal hdc As Long, _
   ByVal X1 As Long, _
   ByVal Y1 As Long, _
   ByVal RadiusX As Long, _
   ByVal RadiusY As Long, _
   ByVal Color As OLE_COLOR, _
   Optional ByVal Thickness As cThickness = Thick)

   Dim Bbg              As Byte
   Dim Gbg              As Byte
   Dim Rbg              As Byte
   Dim savAlpha(1 To 4) As Byte
   Dim Bblend           As Long
   Dim Bgr              As Long
   Dim Cl               As Long
   Dim Gblend           As Long
   Dim Strength         As Long
   Dim StrengthI        As Long
   Dim Quadrant         As Long
   Dim Radius           As Long
   Dim Rblend           As Long
   Dim RX1              As Long
   Dim RX2              As Long
   Dim RY1              As Long
   Dim RY2              As Long
   Dim savX(1 To 4)     As Long
   Dim savY(1 To 4)     As Long
   Dim X4               As Long
   Dim Y4               As Long
   Dim NewColor         As Long
   Dim Ax               As Single
   Dim Ay               As Single
   Dim Bx               As Single
   Dim By               As Single
   Dim L1               As Single
   Dim L2               As Single
   Dim L3               As Single
   Dim L4               As Single
   Dim sngAngle         As Single
   Dim sngPointSpacing  As Single
   Dim X2               As Single
   Dim Xp5              As Single
   Dim Y2               As Single

   m_hDC = hdc

   SetRGBComponents Color

   Radius = RadiusX
   If RadiusY > RadiusX Then
      Radius = RadiusY
   End If

   sngPointSpacing = GetPointSpacing(Radius, Thickness)

   For sngAngle = 0 To HalfPi Step sngPointSpacing
      X2 = RadiusX * Cos(sngAngle)
      Y2 = RadiusY * Sin(sngAngle)
      'Prevents error when vb rounds .5 down
      If X2 = Int(X2) Then X2 = X2 + 0.001
      If Y2 = Int(Y2) Then Y2 = Y2 + 0.001
      For Quadrant = 0 To 3
         Select Case Quadrant
            Case 0 '0-90°
               Ax = X2 + X1 - 0.5
               Ay = -Y2 + Y1 - 0.5
            Case 1 '90-180°
               Ax = X2 + X1 - 0.5
               Ay = Y2 + Y1 - 0.5
            Case 2 '180-270°
               Ax = -X2 + X1 - 0.5
               Ay = Y2 + Y1 - 0.5
            Case 3 '270-360°
               Ax = -X2 + X1 - 0.5
               Ay = -Y2 + Y1 - 0.5
         End Select
         Bx = Ax + 1
         By = Ay + 1
         RX1 = Ax
         RX2 = RX1 + 1
         Xp5 = RX1 + 0.5
         RY1 = Ay
         RY2 = By
         L1 = RY1 + 0.5 - Ay
         L2 = 256 * (Xp5 - Ax) - Xp5 + Ax
         L3 = 255 - L2
         L4 = By - RY2 + 0.5
         savX(1) = RX1
         savY(1) = RY1
         savX(2) = RX2
         savY(2) = RY1
         savY(3) = RY2
         savX(3) = RX1
         savY(4) = RY2
         savX(4) = RX2
         savAlpha(1) = L1 * L2
         savAlpha(2) = L1 * L3
         savAlpha(3) = L4 * L2
         savAlpha(4) = L4 * L3

         For Cl = 1 To 4
            Strength = savAlpha(Cl)
            X4 = savX(Cl)
            Y4 = savY(Cl)
            If Strength > 252 Then '> 99%
               SetPixelV m_hDC, X4, Y4, Color
            Else
               Bgr = GetPixel(m_hDC, X4, Y4)
               If Bgr Then 'if not black
                  Rbg = Bgr And &HFF&
                  Gbg = (Bgr And &HFF00&) \ &H100&
                  Bbg = (Bgr And &HFF0000) \ &H10000
               Else
                  Rbg = 0
                  Gbg = 0
                  Bbg = 0
               End If
               StrengthI = 255 - Strength
               Rblend = StrengthI * Rbg + Strength * Red
               Gblend = StrengthI * Gbg + Strength * Green
               Bblend = StrengthI * Bbg + Strength * Blue
               NewColor = RGB(Rblend \ 256, Gblend \ 256, Bblend \ 256)
               SetPixelV m_hDC, X4, Y4, NewColor
            End If
         Next
      Next
   Next

End Sub

Public Sub ArcGP(ByVal hdc As Long, _
   ByVal X1 As Long, _
   ByVal Y1 As Long, _
   ByVal RadiusX As Long, _
   ByVal RadiusY As Long, _
   ByVal StartAngle As Single, _
   ByVal StopAngle As Single, _
   ByVal Color As OLE_COLOR, _
   Optional ByVal Thickness As cThickness = Thick)

   Dim Bbg              As Byte
   Dim Gbg              As Byte
   Dim Rbg              As Byte
   Dim savAlpha(1 To 4) As Byte
   Dim Bblend           As Long
   Dim Bgr              As Long
   Dim Cl               As Long
   Dim Gblend           As Long
   Dim Strength         As Long
   Dim StrengthI        As Long
   Dim Radius           As Long
   Dim Rblend           As Long
   Dim RX1              As Long
   Dim RX2              As Long
   Dim RY1              As Long
   Dim RY2              As Long
   Dim savX(1 To 4)     As Long
   Dim savY(1 To 4)     As Long
   Dim X4               As Long
   Dim Y4               As Long
   Dim NewColor         As Long
   Dim Ax               As Single
   Dim Ay               As Single
   Dim Bx               As Single
   Dim By               As Single
   Dim L1               As Single
   Dim L2               As Single
   Dim L3               As Single
   Dim L4               As Single
   Dim sngAngle         As Single
   Dim sngPointSpacing  As Single
   Dim X2               As Single
   Dim Xp5              As Single
   Dim Y2               As Single

   m_hDC = hdc

   SetRGBComponents Color

   Radius = RadiusX
   If RadiusY > RadiusX Then
      Radius = RadiusY
   End If

   sngPointSpacing = GetPointSpacing(Radius, Thickness)

   If StartAngle > StopAngle Then
      StopAngle = StopAngle + 360
   End If
   'Convert to Radians
   StartAngle = StartAngle * Rads
   StopAngle = StopAngle * Rads

   For sngAngle = StartAngle To StopAngle Step sngPointSpacing
      X2 = RadiusX * Cos(sngAngle - HalfPi)
      Y2 = RadiusY * Sin(sngAngle - HalfPi)
      'Prevents error when vb rounds .5 down
      If X2 = Int(X2) Then X2 = X2 + 0.001
      If Y2 = Int(Y2) Then Y2 = Y2 + 0.001
      Ax = X2 + X1 - 0.5
      Ay = Y2 + Y1 - 0.5
      Bx = Ax + 1
      By = Ay + 1
      RX1 = Ax
      RX2 = RX1 + 1
      Xp5 = RX1 + 0.5
      RY1 = Ay
      RY2 = By
      L1 = RY1 + 0.5 - Ay
      L2 = 256 * (Xp5 - Ax) - Xp5 + Ax
      L3 = 255 - L2
      L4 = By - RY2 + 0.5
      savX(1) = RX1
      savY(1) = RY1
      savX(2) = RX2
      savY(2) = RY1
      savY(3) = RY2
      savX(3) = RX1
      savY(4) = RY2
      savX(4) = RX2
      savAlpha(1) = L1 * L2
      savAlpha(2) = L1 * L3
      savAlpha(3) = L4 * L2
      savAlpha(4) = L4 * L3

      For Cl = 1 To 4
         Strength = savAlpha(Cl)
         X4 = savX(Cl)
         Y4 = savY(Cl)
         If Strength > 252 Then '> 99%
            SetPixelV m_hDC, X4, Y4, Color
         Else
            Bgr = GetPixel(m_hDC, X4, Y4)
            If Bgr Then 'if not black
               Rbg = Bgr And &HFF&
               Gbg = (Bgr And &HFF00&) \ &H100&
               Bbg = (Bgr And &HFF0000) \ &H10000
            Else
               Rbg = 0
               Gbg = 0
               Bbg = 0
            End If
            StrengthI = 255 - Strength
            Rblend = StrengthI * Rbg + Strength * Red
            Gblend = StrengthI * Gbg + Strength * Green
            Bblend = StrengthI * Bbg + Strength * Blue
            NewColor = RGB(Rblend \ 256, Gblend \ 256, Bblend \ 256)
            SetPixelV m_hDC, X4, Y4, NewColor
         End If
      Next
   Next

End Sub

Private Function GetPointSpacing(Radius As Long, Thickness As cThickness) As Single
   Dim sngLS            As Single

   If Thickness = Thick Then
      sngLS = cThick
   Else
      sngLS = cThin
   End If

   If Radius < 0 Then
      GetPointSpacing = -sngLS / Radius
   ElseIf Radius = 0 Then
      GetPointSpacing = sngLS
   Else
      GetPointSpacing = sngLS / Radius
   End If

End Function

Public Sub CircleDIB(ByVal X1 As Long, _
   ByVal Y1 As Long, _
   ByVal RadiusX As Long, _
   ByVal RadiusY As Long, _
   ByVal Color As OLE_COLOR, _
   Optional ByVal Thickness As cThickness = Thick)

On Error Resume Next
   Dim Bbg              As Byte
   Dim Gbg              As Byte
   Dim Rbg              As Byte
   Dim savAlpha(1 To 4) As Byte
   Dim Cl               As Long
   Dim Strength         As Long
   Dim StrengthI        As Long
   Dim Quadrant         As Long
   Dim Radius           As Long
   Dim RX1              As Long
   Dim RX2              As Long
   Dim RY1              As Long
   Dim RY2              As Long
   Dim savX(1 To 4)     As Long
   Dim savY(1 To 4)     As Long
   Dim X4               As Long
   Dim Y4               As Long
   Dim Ax               As Single
   Dim Ay               As Single
   Dim Bx               As Single
   Dim By               As Single
   Dim L1               As Single
   Dim L2               As Single
   Dim L3               As Single
   Dim L4               As Single
   Dim sngAngle         As Single
   Dim sngPointSpacing  As Single
   Dim X2               As Single
   Dim Xp5              As Single
   Dim Y2               As Single

   If Not InDIBits Then
      MsgBox "You must create a DIB array" & vbCrLf & _
         "before calling CircleDIB."
      Exit Sub
   End If

   SetRGBComponents Color

   Radius = RadiusX
   If RadiusY > RadiusX Then
      Radius = RadiusY
   End If

   sngPointSpacing = GetPointSpacing(Radius, Thickness)

   For sngAngle = 0 To HalfPi Step sngPointSpacing
      X2 = RadiusX * Cos(sngAngle)
      Y2 = RadiusY * Sin(sngAngle)
      'Prevents error when vb rounds .5 down
      If X2 = Int(X2) Then X2 = X2 + 0.001
      If Y2 = Int(Y2) Then Y2 = Y2 + 0.001
      For Quadrant = 0 To 3
         Select Case Quadrant
            Case 0 '0-90°
               Ax = X2 + X1 - 0.5
               Ay = -Y2 + Y1 - 0.5
            Case 1 '90-180°
               Ax = X2 + X1 - 0.5
               Ay = Y2 + Y1 - 0.5
            Case 2 '180-270°
               Ax = -X2 + X1 - 0.5
               Ay = Y2 + Y1 - 0.5
            Case 3 '270-360°
               Ax = -X2 + X1 - 0.5
               Ay = -Y2 + Y1 - 0.5
         End Select

         Bx = Ax + 1
         By = Ay + 1
         RX1 = Ax
         RX2 = RX1 + 1
         Xp5 = RX1 + 0.5
         RY1 = Ay
         RY2 = By
         L1 = RY1 + 0.5 - Ay
         L2 = 256 * (Xp5 - Ax) - Xp5 + Ax
         L3 = 255 - L2
         L4 = By - RY2 + 0.5
         savX(1) = RX1
         savY(1) = RY1
         savX(2) = RX2
         savY(2) = RY1
         savY(3) = RY2
         savX(3) = RX1
         savY(4) = RY2
         savX(4) = RX2
         savAlpha(1) = L1 * L2
         savAlpha(2) = L1 * L3
         savAlpha(3) = L4 * L2
         savAlpha(4) = L4 * L3

         For Cl = 1 To 4
            Strength = savAlpha(Cl)
            X4 = savX(Cl)
            Y4 = savY(Cl)
            'Check if in bounds
            If X4 >= 0 And X4 < m_W1 And Y4 >= 0 And Y4 <= m_H1 Then
               If Strength > 252 Then  '> 99%
                  'No blending
                  buf(X4, Y4).Blue = Blue
                  buf(X4, Y4).Green = Green
                  buf(X4, Y4).Red = Red
               Else
                  Bbg = buf(X4, Y4).Blue
                  Gbg = buf(X4, Y4).Green
                  Rbg = buf(X4, Y4).Red
                  'Blend % of bgColor with % of Color
                  StrengthI = 255 - Strength
                  buf(X4, Y4).Red = (StrengthI * Rbg + Strength * Red) \ 256
                  buf(X4, Y4).Green = (StrengthI * Gbg + Strength * Green) \ 256
                  buf(X4, Y4).Blue = (StrengthI * Bbg + Strength * Blue) \ 256
               End If
            End If
         Next
      Next
   Next
End Sub

Public Sub ArcDIB(ByVal X1 As Long, _
   ByVal Y1 As Long, _
   ByVal RadiusX As Long, _
   ByVal RadiusY As Long, _
   ByVal StartAngle As Single, _
   ByVal StopAngle As Single, _
   ByVal Color As OLE_COLOR, _
   Optional ByVal Thickness As cThickness = Thick)

   Dim Bbg              As Byte
   Dim Gbg              As Byte
   Dim Rbg              As Byte
   Dim savAlpha(1 To 4) As Byte
   Dim Cl               As Long
   Dim Strength         As Long
   Dim StrengthI        As Long
   Dim Radius           As Long
   Dim RX1              As Long
   Dim RX2              As Long
   Dim RY1              As Long
   Dim RY2              As Long
   Dim savX(1 To 4)     As Long
   Dim savY(1 To 4)     As Long
   Dim X4               As Long
   Dim Y4               As Long
   Dim Ax               As Single
   Dim Ay               As Single
   Dim Bx               As Single
   Dim By               As Single
   Dim L1               As Single
   Dim L2               As Single
   Dim L3               As Single
   Dim L4               As Single
   Dim sngAngle         As Single
   Dim sngPointSpacing  As Single
   Dim X2               As Single
   Dim Xp5              As Single
   Dim Y2               As Single

   If Not InDIBits Then
      MsgBox "You must create a DIB array" & vbCrLf & _
         "before calling CircleDIB."
      Exit Sub
   End If

   SetRGBComponents Color

   Radius = RadiusX
   If RadiusY > RadiusX Then
      Radius = RadiusY
   End If

   sngPointSpacing = GetPointSpacing(Radius, Thickness)

   If StartAngle > StopAngle Then
      StopAngle = StopAngle + 360
   End If
   'Convert to Radians
   StartAngle = StartAngle * Rads
   StopAngle = StopAngle * Rads

   For sngAngle = StartAngle To StopAngle Step sngPointSpacing
      X2 = RadiusX * Cos(sngAngle - HalfPi)
      Y2 = RadiusY * Sin(sngAngle + HalfPi)
      'Prevents error when vb rounds .5 down
      If X2 = Int(X2) Then X2 = X2 + 0.001
      If Y2 = Int(Y2) Then Y2 = Y2 + 0.001
      Ax = X2 + X1 - 0.5
      Ay = Y2 + Y1 - 0.5
      Bx = Ax + 1
      By = Ay + 1
      RX1 = Ax
      RX2 = RX1 + 1
      Xp5 = RX1 + 0.5
      RY1 = Ay
      RY2 = By
      L1 = RY1 + 0.5 - Ay
      L2 = 256 * (Xp5 - Ax) - Xp5 + Ax
      L3 = 255 - L2
      L4 = By - RY2 + 0.5
      savX(1) = RX1
      savY(1) = RY1
      savX(2) = RX2
      savY(2) = RY1
      savY(3) = RY2
      savX(3) = RX1
      savY(4) = RY2
      savX(4) = RX2
      savAlpha(1) = L1 * L2
      savAlpha(2) = L1 * L3
      savAlpha(3) = L4 * L2
      savAlpha(4) = L4 * L3

      For Cl = 1 To 4
         Strength = savAlpha(Cl)
         X4 = savX(Cl)
         Y4 = savY(Cl)
         'Check if in bounds
         If X4 >= 0 And X4 < m_W1 And Y4 >= 0 And Y4 <= m_H1 Then
            If Strength > 252 Then  '> 99%
               'No blending
               buf(X4, Y4).Blue = Blue
               buf(X4, Y4).Green = Green
               buf(X4, Y4).Red = Red
            Else
               Bbg = buf(X4, Y4).Blue
               Gbg = buf(X4, Y4).Green
               Rbg = buf(X4, Y4).Red
               'Blend % of bgColor with % of Color
               StrengthI = 255 - Strength
               buf(X4, Y4).Red = (StrengthI * Rbg + Strength * Red) \ 256
               buf(X4, Y4).Green = (StrengthI * Gbg + Strength * Green) \ 256
               buf(X4, Y4).Blue = (StrengthI * Bbg + Strength * Blue) \ 256
            End If
         End If
      Next
   Next
End Sub

Public Sub SetRGBComponents(ByVal Color As OLE_COLOR)

   Color = TranslateColour(Color)
   m_Color = Color 'make available global
   If Color Then
      Red = Color And &HFF&
      Green = Color \ 256 And &HFF
      Blue = Color \ 65536
   Else 'Color is Black
      Red = 0
      Green = 0
      Blue = 0
   End If
End Sub

Public Sub LineGP(ByVal hdc As Long, _
   ByVal X1 As Integer, _
   ByVal Y1 As Integer, _
   ByVal X2 As Integer, _
   ByVal Y2 As Integer, _
   ByVal Color As OLE_COLOR)

   Dim XScope           As Long
   Dim YScope           As Long
   Dim XDir             As Long
   Dim YDir             As Long
   Dim LinearDeviance   As Long
   Dim Counter          As Long
   Dim AntiAliasStrength As Long
   Dim EndPointIntensity As Long

   Const HalfIntensity As Long = 127

   'Blended lines, maximum blend at transition,
   'tapering off too minimum.
   '
   '            ----------
   'Min    Max/Max     Min
   '----------
   '
   m_hDC = hdc
   m_Color = Color

   XScope = X2 - X1
   YScope = Y2 - Y1

   If XScope < 0 Then
      XScope = Abs(XScope)
      XDir = -1
   Else
      XDir = 1
   End If

   If YScope < 0 Then
      YScope = Abs(YScope)
      YDir = -1
   Else
      YDir = 1
   End If

   If XScope + YScope = 0 Then
      'Exit if line length is 0
      Exit Sub
   End If

   SetRGBComponents Color

   If XScope > YScope Then
      'Output EndPoints outside of main loop.
      EndPointIntensity = (85 * YScope) \ XScope
      PutPixelGP X1 - XDir, Y1 - YDir, EndPointIntensity
      PutPixelGP X1 - XDir, Y1, HalfIntensity
      PutPixelGP X2 + XDir, Y2 + YDir, EndPointIntensity
      PutPixelGP X2 + XDir, Y2, HalfIntensity
      '-----
      LinearDeviance = XScope \ 2
      For Counter = 0 To XScope
         'Main line, output full strength direct to hDC.
         SetPixelV m_hDC, X1, Y1, m_Color
         'Output the blended lines for anti-alias effect.
         AntiAliasStrength = (LinearDeviance * 255) \ XScope
         PutPixelGP X1, Y1 - YDir, 255 - AntiAliasStrength
         PutPixelGP X1, Y1 + YDir, AntiAliasStrength
         LinearDeviance = (LinearDeviance + YScope)
         If LinearDeviance >= XScope Then
            LinearDeviance = LinearDeviance - XScope
            Y1 = Y1 + YDir
         End If
         X1 = X1 + XDir
      Next
   Else
      'Output EndPoints outside of main loop.
      EndPointIntensity = (85 * XScope) \ YScope
      PutPixelGP X1 - XDir, Y1 - YDir, EndPointIntensity
      PutPixelGP X1, Y1 - YDir, HalfIntensity
      PutPixelGP X2 + XDir, Y2 + YDir, EndPointIntensity
      PutPixelGP X2, Y2 + YDir, HalfIntensity
      '-----
      LinearDeviance = YScope \ 2
      For Counter = 0 To YScope
         'Main line, output full strength direct to hDC.
         SetPixelV m_hDC, X1, Y1, m_Color
         'Output the blended lines for anti-alias effect.
         AntiAliasStrength = (LinearDeviance * 255) \ YScope
         PutPixelGP X1 - XDir, Y1, 255 - AntiAliasStrength
         PutPixelGP X1 + XDir, Y1, AntiAliasStrength
         LinearDeviance = LinearDeviance + XScope
         If (LinearDeviance >= YScope) Then
            LinearDeviance = LinearDeviance - YScope
            X1 = X1 + XDir
         End If
         Y1 = Y1 + YDir
      Next
   End If

End Sub

Public Sub LineDIB(ByVal X1 As Integer, _
   ByVal Y1 As Integer, _
   ByVal X2 As Integer, _
   ByVal Y2 As Integer, _
   ByVal Color As OLE_COLOR)

   'Blended lines, maximum blend at transition,
   'tapering off too minimum.
   '
   '            ----------
   'Min    Max/Max     Min
   '----------
   '

   If Not InDIBits Then
      MsgBox "You must create a DIB array" & vbCrLf & _
         "before calling LineDIB."
      Exit Sub
   End If

   Dim XScope           As Long
   Dim YScope           As Long
   Dim XDir             As Long
   Dim YDir             As Long
   Dim LinearDeviance   As Long
   Dim Counter          As Long
   Dim AntiAliasStrength As Long
   Dim EndPointIntensity As Long

   Const HalfIntensity As Long = 127

   XScope = X2 - X1
   YScope = Y2 - Y1

   If XScope < 0 Then
      XScope = Abs(XScope)
      XDir = -1
   Else
      XDir = 1
   End If

   If YScope < 0 Then
      YScope = Abs(YScope)
      YDir = -1
   Else
      YDir = 1
   End If

   If XScope + YScope = 0 Then
      'Exit if line length is 0
      Exit Sub
   End If

   SetRGBComponents Color

   If XScope > YScope Then
      'Output EndPoints outside of main loop.
      EndPointIntensity = (85 * YScope) \ XScope
      PutPixelDIB X1 - XDir, Y1 - YDir, EndPointIntensity
      PutPixelDIB X1 - XDir, Y1, HalfIntensity
      PutPixelDIB X2 + XDir, Y2 + YDir, EndPointIntensity
      PutPixelDIB X2 + XDir, Y2, HalfIntensity
      '-----
      LinearDeviance = XScope \ 2
      For Counter = 0 To XScope
         'Main line, output full strength direct to DIB array.
         PutPixelDIB X1, Y1, 255
         'Output the blended lines for anti-alias effect.
         AntiAliasStrength = (LinearDeviance * 255) \ XScope
         PutPixelDIB X1, Y1 - YDir, 255 - AntiAliasStrength
         PutPixelDIB X1, Y1 + YDir, AntiAliasStrength
         LinearDeviance = (LinearDeviance + YScope)
         If LinearDeviance >= XScope Then
            LinearDeviance = LinearDeviance - XScope
            Y1 = Y1 + YDir
         End If
         X1 = X1 + XDir
      Next
   Else
      'Output EndPoints outside of main loop.
      EndPointIntensity = (85 * XScope) \ YScope
      PutPixelDIB X1 - XDir, Y1 - YDir, EndPointIntensity
      PutPixelDIB X1, Y1 - YDir, HalfIntensity
      PutPixelDIB X2 + XDir, Y2 + YDir, EndPointIntensity
      PutPixelDIB X2, Y2 + YDir, HalfIntensity
      '-----
      LinearDeviance = YScope \ 2
      For Counter = 0 To YScope
         'Main line, output full strength direct to DIB array.
         PutPixelDIB X1, Y1, 255
         'Output the blended lines for anti-alias effect.
         AntiAliasStrength = (LinearDeviance * 255) \ YScope
         PutPixelDIB X1 - XDir, Y1, 255 - AntiAliasStrength
         PutPixelDIB X1 + XDir, Y1, AntiAliasStrength
         LinearDeviance = LinearDeviance + XScope
         If (LinearDeviance >= YScope) Then
            LinearDeviance = LinearDeviance - YScope
            X1 = X1 + XDir
         End If
         Y1 = Y1 + YDir
      Next
   End If

End Sub

Public Sub Array2Pic()
   'If we have an array copy back to hDC
   If InDIBits Then
      SetDIBits m_hDC, m_Handle, 0, m_H1, buf(0, 0), Binfo, DIB_RGB_COLORS
      InDIBits = False
      'Erase buf '* Moved to Class_Terminate
   End If
End Sub

Private Sub PutPixelGP(ByVal X As Long, _
   ByVal Y As Long, _
   ByVal Strength As Long)

   Dim Color            As Long
   Dim bgColor          As Long
   Dim Rbg              As Long
   Dim Gbg              As Long
   Dim Bbg              As Long
   Dim Rblend           As Long
   Dim Gblend           As Long
   Dim Bblend           As Long
   Dim StrengthI        As Long

   If Strength > 252 Then '99%
      SetPixelV m_hDC, X, Y, m_Color
   Else
      '##### Get Background Pixel components
      bgColor = GetPixel(m_hDC, X, Y)
      If bgColor Then 'i.e. Not Black
         Rbg = bgColor And &HFF&
         Gbg = (bgColor And &HFF00&) \ &H100&
         Bbg = (bgColor And &HFF0000) \ &H10000
      End If
      '##### Blend % of bgColor with % of m_Color
      StrengthI = 255 - Strength
      Rblend = StrengthI * Rbg + Strength * Red
      Gblend = StrengthI * Gbg + Strength * Green
      Bblend = StrengthI * Bbg + Strength * Blue
      '##### Write
      Color = RGB(Rblend \ 256, Gblend \ 256, Bblend \ 256)
      SetPixelV m_hDC, X, Y, Color
   End If

End Sub

Private Sub PutPixelDIB(ByVal X As Long, _
   ByVal Y As Long, _
   ByVal Strength As Long)

   Dim Rbg              As Long
   Dim Gbg              As Long
   Dim Bbg              As Long
   Dim StrengthI        As Long

   'Check if in bounds
   If X < 0 Or X >= m_W1 Or Y < 0 Or Y > m_H1 Then
      Exit Sub
   End If
   If Strength > 252 Then '99%
      buf(X, Y).Blue = Blue
      buf(X, Y).Green = Green
      buf(X, Y).Red = Red
   Else
      '##### Get Background Pixel components
      Bbg = buf(X, Y).Blue
      Gbg = buf(X, Y).Green
      Rbg = buf(X, Y).Red
      '##### Blend % of bgColor with % of m_Color
      StrengthI = 255 - Strength
      buf(X, Y).Red = (StrengthI * Rbg + Strength * Red) \ 256
      buf(X, Y).Green = (StrengthI * Gbg + Strength * Green) \ 256
      buf(X, Y).Blue = (StrengthI * Bbg + Strength * Blue) \ 256
   End If
End Sub

Private Sub Class_Terminate()
   Erase buf()
End Sub
