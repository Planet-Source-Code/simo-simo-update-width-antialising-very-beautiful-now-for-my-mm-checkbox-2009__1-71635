VERSION 5.00
Begin VB.UserControl mm_checkbox 
   AutoRedraw      =   -1  'True
   ClientHeight    =   930
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1455
   ClipBehavior    =   0  'None
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   62
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   97
   ToolboxBitmap   =   "mon_advanced_checkbox.ctx":0000
   Begin VB.PictureBox picbig 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   435
      Picture         =   "mon_advanced_checkbox.ctx":0312
      ScaleHeight     =   435
      ScaleWidth      =   720
      TabIndex        =   1
      Top             =   75
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox picsmall 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   75
      Picture         =   "mon_advanced_checkbox.ctx":13A4
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   0
      Top             =   555
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image pic_des_big_check 
      Height          =   435
      Left            =   150
      Picture         =   "mon_advanced_checkbox.ctx":181E
      Top             =   1485
      Width           =   720
   End
   Begin VB.Image pic_des_big_uncheck 
      Height          =   435
      Left            =   960
      Picture         =   "mon_advanced_checkbox.ctx":28B0
      Top             =   1500
      Width           =   720
   End
   Begin VB.Image pic_des_small_uncheck 
      Height          =   225
      Left            =   1110
      Picture         =   "mon_advanced_checkbox.ctx":3942
      Top             =   1215
      Width           =   345
   End
   Begin VB.Image pic_des_small_check 
      Height          =   225
      Left            =   345
      Picture         =   "mon_advanced_checkbox.ctx":3DBC
      Top             =   1185
      Width           =   345
   End
End
Attribute VB_Name = "mm_checkbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'EVENTS.
Public Event Click()
Public Event DoubleClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseEnters(ByVal X As Long, ByVal Y As Long)
Public Event MouseLeaves(ByVal X As Long, ByVal Y As Long)


Private udtPoint As POINTAPI
Private bolMouseDown As Boolean
Private bolMouseOver As Boolean
'Private bolHasFocus As Boolean
Private bolEnabled As Boolean
Private bolChecked As Boolean
Private bolSmall As Boolean
Private lonRoundValue As Long
Private lonRect As Long
Private button_clique As Integer

Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long


'to draw radial cercle
Dim AA1 As New LineGS 'DrawRadial

Private m_Activecolor As OLE_COLOR
Private m_desActivecolor As OLE_COLOR

'Private Declare Function CreateEllipticRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function FrameRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Private Declare Function FillRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
'Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
'Private Declare Function Ellipse Lib "gdi32" _
    (ByVal hdc As Long, ByVal X1 As Long, _
    ByVal Y1 As Long, ByVal X2 As Long, _
    ByVal Y2 As Long) As Long

Sub mon_gradient(mcolor As Long, X As Integer, Y As Integer, iCircle As Integer)
   Dim I As Integer
       
   UserControl.Cls
   UserControl.DrawStyle = 5
   UserControl.FillStyle = 0
                                                      '|                                            |

    UserControl.Cls
    
      With UserControl
         'Copy DIBits to array
         AA1.DIB .hdc, .Image.Handle, .ScaleWidth, .ScaleHeight
      End With
   
   '1st circle
    If Not Small Then
        For I = 5 To 6
            AA1.CircleDIB X, Y, iCircle + I, iCircle + I, &HDAD4CE  'RGB(128, 128, 128) 'vbRed ''100, 100, I * 0.75, I * 0.75, vbRed
        Next I
    Else
            AA1.CircleDIB X, Y, iCircle + 3, iCircle + 3, &HDAD4CE  'RGB(128, 128, 128) 'vbRed ''100, 100, I * 0.75, I * 0.75, vbRed
    End If
   
'    'simulate a blendcolor circle
    AA1.CircleDIB X, Y, iCircle + 1, iCircle + 1, BlendColor(mcolor, vbWhite, 100) '&HDAD4CE    'RGB(128, 128, 128) 'vbRed ''100, 100, I * 0.75, I * 0.75, vbRed
    AA1.CircleDIB X, Y, iCircle + 2, iCircle + 2, BlendColor(mcolor, vbWhite, 50) '&HDAD4CE    'RGB(128, 128, 128) 'vbRed ''100, 100, I * 0.75, I * 0.75, vbRed
   
   
     For I = iCircle To 0 Step -1
        AA1.CircleDIB X, Y, I, I, BlendColor(mcolor, vbWhite, I * (255 / iCircle))
     Next I
     
     'refresh picture in usercontrol
      AA1.Array2Pic
      
   
End Sub

Public Sub About()
Attribute About.VB_UserMemId = -552
    dlgAbout.Show 1
End Sub


Private Function PointInControl(X As Single, Y As Single) As Boolean
  If X >= 0 And X <= UserControl.ScaleWidth And _
    Y >= 0 And Y <= UserControl.ScaleHeight Then
    PointInControl = True
  End If
End Function

Private Sub PaintControl()
    
    
    UserControl.Refresh
    UserControl.Picture = LoadPicture("")
    UserControl.Refresh
    
    If Small Then
'        Round1 = 10
        UserControl.Width = (picsmall.Width + 1) * Screen.TwipsPerPixelX
        UserControl.Height = (picsmall.Height + 1) * Screen.TwipsPerPixelY
        UserControl.Picture = picsmall.Picture
        If Checked Then
            mon_gradient m_Activecolor, 7, (UserControl.ScaleHeight / 2) - 1, 3 '6 '8  'UserControl.ScaleHeight / 2 - 1, 9
        Else
            mon_gradient m_desActivecolor, UserControl.ScaleWidth - 9, (UserControl.ScaleHeight / 2) - 1, 3 '6
        End If
    Else
'        Round1 = 26
        UserControl.Width = (picbig.Width + 1) * Screen.TwipsPerPixelX
        UserControl.Height = (picbig.Height + 1) * Screen.TwipsPerPixelY
        UserControl.Picture = picbig.Picture
        If Checked Then
            mon_gradient m_Activecolor, 15, (UserControl.ScaleHeight / 2) - 1, 8 '9 '8  'UserControl.ScaleHeight / 2 - 1, 9
        Else
            mon_gradient m_desActivecolor, UserControl.ScaleWidth - 16, (UserControl.ScaleHeight / 2) - 1, 8 '9
        End If
    End If
    
End Sub

Public Property Get Activecolor() As OLE_COLOR
   Activecolor = m_Activecolor
End Property
Public Property Get desActivecolor() As OLE_COLOR
   desActivecolor = m_desActivecolor
End Property
Public Property Let Activecolor(ByVal New_Activecolor As OLE_COLOR)
   m_Activecolor = New_Activecolor
   PropertyChanged "Activecolor"
   PaintControl
End Property
Public Property Let desActivecolor(ByVal New_desActivecolor As OLE_COLOR)
   m_desActivecolor = New_desActivecolor
   PropertyChanged "desActivecolor"
   PaintControl
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Button Enabled/Disable."
Enabled = bolEnabled
End Property

Public Property Get Small() As Boolean
Small = bolSmall
End Property
Public Property Get Checked() As Boolean
Checked = bolChecked
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
bolEnabled = NewValue
PropertyChanged "Enabled"



If bolEnabled Then

    pic_des_small_check.Top = -200
    pic_des_small_uncheck.Top = -200
    
    pic_des_big_check.Top = -200
    pic_des_big_uncheck.Left = -200

Else
    If bolSmall Then
        If Checked Then
            pic_des_small_check.Top = 0
            pic_des_small_check.Left = 0
        Else
            pic_des_small_uncheck.Top = 0
            pic_des_small_uncheck.Left = 0
        End If
    
    Else
        If Checked Then
            pic_des_big_check.Top = 0
            pic_des_big_check.Left = 0
        Else
            pic_des_big_uncheck.Top = 0
            pic_des_big_uncheck.Left = 0
        End If
    End If
End If


UserControl.Enabled = bolEnabled

'PaintControl
End Property

Public Property Let Small(ByVal NewValue As Boolean)
bolSmall = NewValue
PropertyChanged "Small"

PaintControl

If Small = True Then
    RoundedValue = 10
Else
    RoundedValue = 26
End If


End Property
Public Property Let Checked(ByVal NewValue As Boolean)
bolChecked = NewValue
PropertyChanged "Checked"

PaintControl

Enabled = Not Enabled
Enabled = Not Enabled

End Property
Public Property Get RoundedValue() As Long
Attribute RoundedValue.VB_Description = "Button Border Rounded Value."
RoundedValue = lonRoundValue
End Property

Public Property Let RoundedValue(ByVal NewValue As Long)
lonRoundValue = NewValue
PropertyChanged "RoundedValue"
PaintControl

lonRect = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, lonRoundValue, lonRoundValue)     '- 1
SetWindowRgn UserControl.hWnd, lonRect, True

'UserControl.ForeColor = vbBlack
'RoundRect UserControl.hdc, 0, 0, ScaleWidth - 1, ScaleHeight - 1, lonRoundValue, lonRoundValue 'lonRoundValue, lonRoundValue

End Property

Private Sub UserControl_Click()
If bolEnabled = True Then
    If button_clique = 1 Then
        
        Checked = Not Checked
        PaintControl
        
        RaiseEvent Click
        RaiseEvent MouseLeaves(0, 0)
    End If
End If
End Sub

Private Sub UserControl_Initialize()
m_Activecolor = &H8000&
m_desActivecolor = &H808080

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bolEnabled = True Then
    button_clique = Button
    If Button = 1 Then
        bolMouseDown = True
        RaiseEvent MouseDown(Button, Shift, X, Y)
'        PaintControl
    End If
End If

End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bolEnabled = False Then Exit Sub
    RaiseEvent MouseMove(Button, Shift, X, Y)
    SetCapture hWnd
    If PointInControl(X, Y) Then
        'pointer on control
        If Not bolMouseOver Then
            bolMouseOver = True
            RaiseEvent MouseEnters(udtPoint.X, udtPoint.Y)
        End If
    Else
        'pointer out of control
        bolMouseOver = False
        bolMouseDown = False
        ReleaseCapture
        RaiseEvent MouseLeaves(udtPoint.X, udtPoint.Y)
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bolEnabled = True Then
    button_clique = Button
    If Button = 1 Then
        RaiseEvent MouseUp(Button, Shift, X, Y)
        bolMouseDown = False
    End If
End If
End Sub

Private Sub UserControl_Paint()
PaintControl
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'On Error Resume Next
With PropBag
    
    Let Enabled = .ReadProperty("Enabled", True)
    Let Checked = .ReadProperty("Checked", False)
    Let Small = .ReadProperty("Small", True)
    Let RoundedValue = .ReadProperty("RoundedValue", 5)
    Let Activecolor = .ReadProperty("Activecolor", m_Activecolor) ' &H117B28) 'vbGreen)
    Let desActivecolor = .ReadProperty("desActivecolor", m_desActivecolor) ' &H117B28) 'vbGreen)
End With
End Sub
Private Sub UserControl_Resize()
    
    If Small Then
        UserControl.Width = (picsmall.Width + 1) * Screen.TwipsPerPixelX
        UserControl.Height = (picsmall.Height + 1) * Screen.TwipsPerPixelY
    Else
        UserControl.Width = (picbig.Width + 1) * Screen.TwipsPerPixelX
        UserControl.Height = (picbig.Height + 1) * Screen.TwipsPerPixelY
    End If
    
End Sub
Private Sub UserControl_Terminate()
bolMouseDown = False
bolMouseOver = False
'bolHasFocus = False
'UserControl.Cls
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'On Error Resume Next
With PropBag
    .WriteProperty "Enabled", bolEnabled, True
    .WriteProperty "Checked", bolChecked, False
    .WriteProperty "Small", bolSmall, True
    .WriteProperty "RoundedValue", lonRoundValue, 5
    .WriteProperty "Activecolor", m_Activecolor, &H8000& '&H117B28 'vbGreen
    .WriteProperty "desActivecolor", m_desActivecolor, &H808080 '&H94A392
End With
End Sub
Private Sub UserControl_InitProperties()
Let Enabled = True
Let Checked = False
Let Small = False 'True
Let RoundedValue = 26 '5

m_Activecolor = &H8000&
m_desActivecolor = &H808080

End Sub
