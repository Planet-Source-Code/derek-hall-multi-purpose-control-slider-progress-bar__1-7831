VERSION 5.00
Begin VB.UserControl ProgressBar 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2265
   EditAtDesignTime=   -1  'True
   PropertyPages   =   "ProgressBar.ctx":0000
   ScaleHeight     =   29
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   151
   Begin VB.PictureBox Tiles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      Picture         =   "ProgressBar.ctx":0014
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   48
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label txt_Value 
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   570
   End
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const SRCCOPY = &HCC0020     ' dest = source
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private BarLength As Long
Const ButtonWidth = 10
Event Click()

Private Type ColourData
  R As Byte
  G As Byte
  B As Byte
End Type

Public Enum YesNOData
  No = 0
  yes = 1
End Enum

Public Enum BarTypeData
  Standard = 0
  Custom = 1
  Random = 2
  HotStuff = 3
  Level = 4
  IceCool = 5
  WhiteHot = 6
  DialType = 7
End Enum

Public Enum BarStyle
  Flat = 0
  Raised = 1
  Sunk = 2
End Enum
 
Public Enum BarType
  Std = 0
  Precise = 1
  Progress = 2
End Enum

Private BarControl As BarType
Private Const New_BarControl = Precise

Private BarValue As Long
Private Const New_BarValue = 0

Private ColourScheme As BarTypeData
Private Const New_ColourScheme = Gray

Private BarMaxValue As Integer
Private Const New_BarMaxValue = 32767

Private BarMinValue As Integer
Private Const New_BarMinValue = 0

Private Bar3D As BarStyle
Private Const New_Bar3D = Sunk

Private BarSChange As Integer
Private Const New_BarSChange = 1

Private BarBChange As Integer
Private Const New_BarBChange = 5

Private BarPreValue As Integer
Private Const New_BarPreValue = 0

Private ShadingDirection As YesNOData
Private Const New_ShadingDirection = No

Private BarShowValues As YesNOData
Private Const New_BarShowValues = yes


Private ButtonsAreOn As YesNOData
Private Const New_ButtonsAreOn = True

Private BarBackColor As Long
Private Const New_BarBackColor = 14737632


Private RgbColour As ColourData
Private CurrentPosition As Integer
Private CurrentDivider As Single
Private MouseIsDown As Byte
Private HighestValue As Long
Private k As Single
Private ColorDivision As Single

Private PreCheckColor As Boolean
Private CheckedTheColor As ColourData
'

'***********

Private Sub DisplayUp()
txt_Value.Visible = False
  Dim i As Single, j As Single, M As Integer
  k = (BarLength) / 2
  ColorDivision = 255 / k
  Select Case BarStyle
    Case Standard
      k = (BarLength)
      ColorDivision = 255 / k
      For i = 0 To k * 1
        MixerF i, 4, 0, 4, 0, 4, 0: DrawF i, 0
      Next i
    Case Custom, Random
      k = (BarLength)
      ColorDivision = 255 / k
      For i = 0 To k * 1
        If ShadingDirection Then
          MixerF i, 3, 0, 3, 0, 3, 0: DrawR i, 0
        Else
          MixerF i, 3, 0, 3, 0, 3, 0: DrawF i, 0
        End If
      Next i
    Case HotStuff
      For i = 0 To k * 1
        If ShadingDirection Then
          MixerF i, 1, 255, 0, 0, 0, 0: DrawR i, 0
          MixerF i, 0, 255, 2, 255, 0, 0: DrawF i, 0
        Else
          MixerF i, 1, 255, 0, 0, 0, 0: DrawF i, 0
          MixerF i, 0, 255, 2, 255, 0, 0: DrawR i, 0
        End If
      Next i

    Case WhiteHot
      For i = 0 To k * 1
        If ShadingDirection Then
          MixerF i, 0, 255, 1, 255, 0, 0: DrawR i, 0
          MixerF i, 0, 255, 0, 255, 2, 255: DrawF i, 0 'WhiteHot
        Else
          MixerF i, 0, 255, 1, 255, 0, 0: DrawF i, 0
          MixerF i, 0, 255, 0, 255, 2, 255: DrawR i, 0 'WhiteHot
        End If
      Next i
      
    Case IceCool
      For i = 0 To k * 1
        If ShadingDirection Then
          MixerF i, 0, 0, 0, 0, 1, 255: DrawR i, 0
          MixerF i, 2, 255, 2, 255, 0, 255: DrawF i, 0
        Else
          MixerF i, 2, 255, 2, 255, 0, 255: DrawR i, 0
          MixerF i, 0, 0, 0, 0, 1, 255: DrawF i, 0
        End If
      Next i
    
    Case Level
      For i = 0 To k * 1
        If ShadingDirection Then
          MixerF i, 1, 0, 0, 255, 0, 0: DrawF i, 0
          MixerF i, 0, 255, 1, 0, 0, 0: DrawR i, 0
        Else
          MixerF i, 0, 255, 1, 0, 0, 0: DrawF i, 0
          MixerF i, 1, 0, 0, 255, 0, 0: DrawR i, 0
        End If
      Next i
    Case DialType
      k = (BarLength) / 4
      ColorDivision = 255 / k
      For i = 0 To k * 1
        MixerF i, 0, 255, 1, 255, 0, 0: DrawF i, 0
        MixerF i, 0, 255, 1, 255, 0, 0: DrawR i, 0
        MixerF i, 2, 255, 0, 255, 1, 255: DrawF i, 1
        MixerF i, 2, 255, 0, 255, 1, 255: DrawR i, 1
          'MixerF i, 0, 50, 1, 255, 0, 255: DrawF i, 0
          'MixerF i, 0, 50, 1, 255, 0, 255: DrawR i, 0
          'MixerF i, 2, 255, 0, 255, 0, 255: DrawF i, 1
          'MixerF i, 2, 255, 0, 255, 0, 255: DrawR i, 1
      Next i

  End Select
  
  
  On Error Resume Next
  HighestValue = -(BarMinValue)
  HighestValue = HighestValue + BarMaxValue
  CurrentPosition = (BarLength / HighestValue) * (BarValue - BarMinValue)
  CurrentPosition = CurrentPosition + ButtonWidth
  CurrentDivider = (BarLength / (HighestValue))
  
  If Not (BarMaxValue = BarMinValue) Then
  If BarShowValues = yes Then
    If CurrentPosition < (BarLength / 2) + 10 Then
      txt_Value.Left = (CurrentPosition) + 8
      txt_Value.Alignment = 0
    Else
      txt_Value.Left = (CurrentPosition) - 46
      txt_Value.Alignment = 1
    End If
    If CheckColorNow(UserControl.Point(CurrentPosition, 5)) Or CheckColorNow(UserControl.Point(CurrentPosition - 1, 5)) Then
      txt_Value.ForeColor = 0 '16777215 'white
    Else
      txt_Value.ForeColor = 16777215 'white
    End If
    txt_Value = Me.BarCurrentValue
  End If
  
  Select Case BarControl
    Case Precise
      If CurrentPosition > ButtonWidth + 4 Then Line (CurrentPosition - 4, 1)-(CurrentPosition - 4, 13), RGB(255, 255, 255)
      If CurrentPosition > ButtonWidth + 3 Then Line (CurrentPosition - 3, 1)-(CurrentPosition - 3, 13), RGB(190, 190, 190)
      If CurrentPosition > ButtonWidth + 2 Then Line (CurrentPosition - 2, 1)-(CurrentPosition - 2, 13), RGB(190, 190, 190)
      If CurrentPosition > ButtonWidth + 1 Then Line (CurrentPosition - 1, 1)-(CurrentPosition - 1, 13), RGB(0, 0, 0)
      If CurrentPosition < ButtonWidth + BarLength - 1 Then Line (CurrentPosition + 1, 1)-(CurrentPosition + 1, 13), RGB(255, 255, 255)
      If CurrentPosition < ButtonWidth + BarLength - 2 Then Line (CurrentPosition + 2, 1)-(CurrentPosition + 2, 13), RGB(190, 190, 190)
      If CurrentPosition < ButtonWidth + BarLength - 3 Then Line (CurrentPosition + 3, 1)-(CurrentPosition + 3, 13), RGB(190, 190, 190)
      If CurrentPosition < ButtonWidth + BarLength - 4 Then Line (CurrentPosition + 4, 1)-(CurrentPosition + 4, 13), RGB(0, 0, 0)
    
    Case Std
      If CurrentPosition > ButtonWidth + 4 Then Line (CurrentPosition - 4, 1)-(CurrentPosition - 4, 13), RGB(255, 255, 255)
      If CurrentPosition > ButtonWidth + 3 Then Line (CurrentPosition - 3, 1)-(CurrentPosition - 3, 13), RGB(255, 255, 255)
      If CurrentPosition > ButtonWidth + 2 Then Line (CurrentPosition - 2, 1)-(CurrentPosition - 2, 13), RGB(190, 190, 190)
      If CurrentPosition > ButtonWidth + 1 Then Line (CurrentPosition - 1, 1)-(CurrentPosition - 1, 13), RGB(190, 190, 190)
      Line (CurrentPosition, 1)-(CurrentPosition, 13), RGB(190, 190, 190)
      If CurrentPosition < ButtonWidth + BarLength - 1 Then Line (CurrentPosition + 1, 1)-(CurrentPosition + 1, 13), RGB(190, 190, 190)
      If CurrentPosition < ButtonWidth + BarLength - 2 Then Line (CurrentPosition + 2, 1)-(CurrentPosition + 2, 13), RGB(190, 190, 190)
      If CurrentPosition < ButtonWidth + BarLength - 3 Then Line (CurrentPosition + 3, 1)-(CurrentPosition + 3, 13), RGB(0, 0, 0)
      If CurrentPosition < ButtonWidth + BarLength - 4 Then Line (CurrentPosition + 4, 1)-(CurrentPosition + 4, 13), RGB(0, 0, 0)
    
    Case Progress
      ButtonsAreOn = No
      If BarReverseShade Then

        For M = 10 To (((BarLength) / BarMaxValue) * (BarValue) + 10) '- 1
          Line (M, 1)-(M, 13), RGB(140, 140, 140)
        Next M
      Else
        For M = CurrentPosition To BarLength + 10
          Line (M, 1)-(M, 13), RGB(140, 140, 140)
        Next M
      End If
      If Not CurrentPosition = 10 Then Line (CurrentPosition - 1, 1)-(CurrentPosition - 1, 13), RGB(255, 255, 255)
      If Not CurrentPosition = BarLength + 10 Then Line (CurrentPosition, 1)-(CurrentPosition, 13), RGB(0, 0, 0)
  
  End Select
  End If
  Select Case Bar3D
    Case Raised
      
      Line (ButtonWidth, 0)-(BarLength + ButtonWidth + 2, 0), RGB(230, 230, 230)
      Line (ButtonWidth, 13)-(BarLength + ButtonWidth + 2, 13), RGB(0, 0, 0)
      Line (ButtonWidth, 0)-(ButtonWidth, 13), RGB(230, 230, 230)
      Line (BarLength + ButtonWidth + 1, 0)-(BarLength + ButtonWidth + 1, 13), RGB(0, 0, 0)
      
    Case Sunk
      Line (ButtonWidth, 0)-(BarLength + ButtonWidth + 2, 0), RGB(0, 0, 0)
      Line (ButtonWidth, 13)-(BarLength + ButtonWidth + 2, 13), RGB(230, 230, 230)
      Line (ButtonWidth, 0)-(ButtonWidth, 13), RGB(0, 0, 0)
      Line (BarLength + ButtonWidth + 1, 0)-(BarLength + ButtonWidth + 1, 13), RGB(230, 230, 230)
  End Select
'****Show Value
  If BarShowValues = yes Then
    txt_Value.Visible = True
  End If
'****End of Show Value
  

  If ButtonsAreOn Then
    If MouseIsDown = 1 Then
      BitBlt UserControl.hDC, 0, 0, 10, 14, Tiles.hDC, 0, 14, SRCCOPY
    Else
      BitBlt UserControl.hDC, 0, 0, 10, 14, Tiles.hDC, 0, 0, SRCCOPY
    End If
    If MouseIsDown = 2 Then
      BitBlt UserControl.hDC, BarLength + ButtonWidth, 0, 10, 14, Tiles.hDC, 10, 14, SRCCOPY
    Else
      BitBlt UserControl.hDC, BarLength + ButtonWidth, 0, 10, 14, Tiles.hDC, 10, 0, SRCCOPY
    End If
  End If

  UserControl.Refresh
End Sub

Private Sub txt_Value_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  UserControl_MouseDown Button, Shift, txt_Value.Left + (X / 15), Y
End Sub

Private Sub txt_Value_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button <> 0 Then UserControl_MouseMove Button, Shift, txt_Value.Left + (X / 15), Y
End Sub


Private Sub UserControl_Initialize()
UserControl_Resize
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If BarControl = Progress Or BarMaxValue = BarMinValue Then Exit Sub
  On Error Resume Next
  
  If X > 10 And X < (UserControl.ScaleWidth - 10) Then
    If Button = 2 Then
      BarCurrentValue = BarMinValue + ((X - 10) / CurrentDivider)
      MouseIsDown = 0
      Exit Sub
    Else
      If Not (X > (CurrentPosition - 6) And X < (CurrentPosition + 6)) Then
        If (X - 10) < CurrentPosition Then
          BarCurrentValue = BarCurrentValue - BarBChange
        Else
          BarCurrentValue = BarCurrentValue + BarBChange
        End If
      Else
        BarCurrentValue = BarMinValue + ((X - 10) / CurrentDivider)
        MouseIsDown = 0
      End If
    End If
  Else
    If X < 10 Then
      MouseIsDown = 1
      BarCurrentValue = BarCurrentValue - BarSChange
    Else
      If X > (UserControl.ScaleWidth - 10) Then
        MouseIsDown = 2
        BarCurrentValue = BarCurrentValue + BarSChange
      End If
    End If
  End If
  
  DisplayUp
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If BarControl = Progress Or BarMaxValue = BarMinValue Then Exit Sub
  If X < 10 Or X > (UserControl.ScaleWidth - 10) Then Exit Sub
  If Button = 1 Then
    If X > (CurrentPosition - 6) And X < (CurrentPosition + 6) Then MouseIsDown = 3
  Else
    If Button = 2 Then BarCurrentValue = BarMinValue + ((X - 10) / CurrentDivider)
  End If
  If MouseIsDown = 3 Then BarCurrentValue = BarMinValue + ((X - 10) / CurrentDivider)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If BarControl = Progress Or BarMaxValue = BarMinValue Then Exit Sub
  If ButtonsAreOn Then
    BitBlt UserControl.hDC, 0, 0, 10, 14, Tiles.hDC, 0, 0, SRCCOPY
    BitBlt UserControl.hDC, BarLength + ButtonWidth, 0, 10, 14, Tiles.hDC, 10, 0, SRCCOPY
  End If
  MouseIsDown = 0
  UserControl.Refresh
End Sub

Private Sub UserControl_Paint()
  DisplayUp
End Sub

Private Sub UserControl_Resize()
  UserControl.ScaleHeight = 14
  If (UserControl.ScaleWidth) < (ButtonWidth * 2) Then UserControl.ScaleWidth = (ButtonWidth * 2) + 5
  BarLength = (UserControl.ScaleWidth - (ButtonWidth * 2)) '+ 1
  DisplayUp
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  ColourScheme = PropBag.ReadProperty("ColourScheme", New_ColourScheme)
  BarValue = PropBag.ReadProperty("BarValue", New_BarValue)
  BarMaxValue = PropBag.ReadProperty("BarMaxValue", New_BarMaxValue)
  BarMinValue = PropBag.ReadProperty("BarMinValue", New_BarMinValue)
  Bar3D = PropBag.ReadProperty("Bar3D", New_Bar3D)
  BarSChange = PropBag.ReadProperty("BarSChange", New_BarSChange)
  BarBChange = PropBag.ReadProperty("BarBChange", New_BarBChange)
  BarPreValue = PropBag.ReadProperty("BarPreValue", New_BarPreValue)
  ButtonsAreOn = PropBag.ReadProperty("ButtonsAreOn", New_ButtonsAreOn)
  UserControl.ForeColor = PropBag.ReadProperty("BarBackColor", New_BarBackColor)
  ShadingDirection = PropBag.ReadProperty("ShadingDirection", New_ShadingDirection)
  BarShowValues = PropBag.ReadProperty("BarShowValues", New_BarShowValues)
  BarControl = PropBag.ReadProperty("BarControl", New_BarControl)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "ColourScheme", ColourScheme, New_ColourScheme
  PropBag.WriteProperty "BarValue", BarValue, New_BarValue
  PropBag.WriteProperty "BarMaxValue", BarMaxValue, New_BarMaxValue
  PropBag.WriteProperty "BarMinValue", BarMinValue, New_BarMinValue
  PropBag.WriteProperty "Bar3D", Bar3D, New_Bar3D
  PropBag.WriteProperty "BarSChange", BarSChange, New_BarSChange
  PropBag.WriteProperty "BarBChange", BarBChange, New_BarBChange
  PropBag.WriteProperty "BarPreValue", BarPreValue, New_BarPreValue
  PropBag.WriteProperty "ButtonsAreOn", ButtonsAreOn, New_ButtonsAreOn
  PropBag.WriteProperty "BarBackColor", UserControl.ForeColor, New_BarBackColor
  PropBag.WriteProperty "ShadingDirection", ShadingDirection, New_ShadingDirection
  PropBag.WriteProperty "BarShowValues", BarShowValues, New_BarShowValues
  PropBag.WriteProperty "BarControl", BarControl, New_BarControl
End Sub

Public Property Get BarStyle() As BarTypeData
  BarStyle = ColourScheme
End Property
Public Property Let BarStyle(ByVal New_Value As BarTypeData)
  ColourScheme = New_Value
  PropertyChanged "ColourScheme"
  DisplayUp
  If BarStyle = Random Then BarCustomColor = (16777215 * Rnd)
End Property

Public Property Get BarCurrentValue() As Long
  BarCurrentValue = BarValue
End Property

Public Property Let BarCurrentValue(ByVal New_Value As Long)
  BarPreviousValue = BarCurrentValue
  If New_Value < BarMinValue Then
    BarValue = BarMinValue
  Else
    If New_Value > BarMaxValue Then
      BarValue = BarMaxValue
    Else
      BarValue = New_Value
    End If
  End If
  PropertyChanged "BarValue"
  DisplayUp
  RaiseEvent Click
End Property

Public Property Get BarCVMax() As Long
  BarCVMax = BarMaxValue
End Property
Public Property Let BarCVMax(ByVal New_Value As Long)
  BarMaxValue = New_Value
  PropertyChanged "BarMaxValue"
  BarValue = BarValue
  If BarCurrentValue > BarMaxValue Then BarCurrentValue = BarMaxValue
End Property
Public Property Get BarCVMin() As Long
  BarCVMin = BarMinValue
End Property
Public Property Let BarCVMin(ByVal New_Value As Long)
  BarMinValue = New_Value
  PropertyChanged "BarMinValue"
  BarValue = BarValue
  If BarCurrentValue < BarMinValue Then BarCurrentValue = BarMinValue
End Property

Public Property Get BarAppearance() As BarStyle
  BarAppearance = Bar3D
End Property
Public Property Let BarAppearance(ByVal New_Value As BarStyle)
  Bar3D = New_Value
  PropertyChanged "Bar3D"
  DisplayUp
End Property

Public Property Get BarSmallChange() As Integer
  BarSmallChange = BarSChange
End Property
Public Property Let BarSmallChange(ByVal New_Value As Integer)
  BarSChange = New_Value
  PropertyChanged "BarSChange"
End Property

Public Property Get BarBigChange() As Integer
  BarBigChange = BarBChange
End Property
Public Property Let BarBigChange(ByVal New_Value As Integer)
  BarBChange = New_Value
  PropertyChanged "BarBChange"
End Property

Public Property Get BarPreviousValue() As Integer
  BarPreviousValue = BarPreValue
End Property
Public Property Let BarPreviousValue(ByVal New_Value As Integer)
  BarPreValue = New_Value
  PropertyChanged "BarPreValue"
End Property

Public Property Get BarShowButtons() As YesNOData
  BarShowButtons = ButtonsAreOn
End Property
Public Property Let BarShowButtons(ByVal New_Value As YesNOData)
  ButtonsAreOn = New_Value
  PropertyChanged "ButtonsAreOn"
  DisplayUp
End Property

Public Property Get BarReverseShade() As YesNOData
  BarReverseShade = ShadingDirection
End Property

Public Property Let BarReverseShade(ByVal New_Value As YesNOData)
  ShadingDirection = New_Value
  PropertyChanged "ShadingDirection"
  DisplayUp
End Property

Public Property Get BarShowValue() As YesNOData
  BarShowValue = BarShowValues
End Property

Public Property Let BarShowValue(ByVal New_Value As YesNOData)

  BarShowValues = New_Value
  PropertyChanged "BarShowValues"
  If BarShowValue Then
    txt_Value.Visible = True
  Else
    txt_Value.Visible = False
  End If
  DisplayUp
End Property

Public Property Get BarSlider() As BarType
  BarSlider = BarControl
End Property

Public Property Let BarSlider(ByVal New_Value As BarType)
  BarControl = New_Value
  PropertyChanged "BarControl"
  DisplayUp
End Property


Public Property Get BarCustomColor() As OLE_COLOR
  BarCustomColor = UserControl.ForeColor()
End Property

Public Property Let BarCustomColor(ByVal New_Value1 As OLE_COLOR)
Dim New_Value As Long
On Error GoTo ExitIt
UserControl.ForeColor() = New_Value1
New_Value = UserControl.ForeColor()
  If New_Value < 0 Then
    BarBackColor = 0
  Else
    If New_Value > 16777215 Then
      BarBackColor = 16777215
    Else
      BarBackColor = New_Value
    End If
  End If
    PropertyChanged "BarBackColor"
    DisplayUp
ExitIt:
End Property

Private Sub Reds(ByVal i As Long)
  RgbColour.R = ColorDivision * i
  RgbColour.G = 0
  RgbColour.B = 0
End Sub

Private Sub Oranges(ByVal i As Long)
  RgbColour.R = 255
  RgbColour.G = (ColorDivision) * i
  RgbColour.B = 0
End Sub

Private Sub Greens(ByVal i As Long)
  RgbColour.R = 255 - (ColorDivision) * i
  RgbColour.G = 255 - (ColorDivision) * (i / 4)
  RgbColour.B = 0
End Sub

Private Sub Mixer1(ByVal i As Long)
  RgbColour.R = 0
  RgbColour.G = 0
  RgbColour.B = ColorDivision * i
End Sub
Private Sub Mixer2(ByVal i As Long)
  RgbColour.R = 255 - (ColorDivision) * i
  RgbColour.G = 255 - (ColorDivision) * i
  RgbColour.B = 255
End Sub

Private Sub MixerF(ByVal i As Long, ROn As Byte, RValue As Byte, GOn As Byte, GValue As Byte, BOn As Byte, BValue As Byte)
On Error Resume Next
With RgbColour
  Select Case ROn
    Case 0
      .R = RValue
    Case 1
      .R = (ColorDivision) * i
    Case 2
      .R = RValue - (ColorDivision) * i
    Case 3
      .R = (((BarCustomColor Mod 65536) Mod 256) / BarLength) * i
    Case 4
      .R = ((BarCustomColor Mod 65536) Mod 256)

  End Select
  
  Select Case GOn
    Case 0
      .G = GValue
    Case 1
      .G = (ColorDivision) * i
    Case 2
      .G = GValue - (ColorDivision) * i
    Case 3
      .G = (((BarCustomColor Mod 65536) \ 256) / BarLength) * i
    Case 4
      .G = ((BarCustomColor Mod 65536) \ 256)
    End Select
  
  Select Case BOn
    Case 0
      .B = BValue
    Case 1
      .B = (ColorDivision) * i
    Case 2
      .B = BValue - (ColorDivision) * i
    Case 3
      .B = ((BarCustomColor \ 65536) / BarLength) * i
    Case 4
      .B = (BarCustomColor \ 65536)
  End Select
End With
End Sub
Private Sub DrawF(ByVal i As Long, TimesItBy As Integer)
  If Bar3D = Flat Then
    Line (i + (k * TimesItBy) + ButtonWidth, 0)-(i + (k * TimesItBy) + ButtonWidth, 14), RGB(RgbColour.R, RgbColour.G, RgbColour.B)
  Else
    Line (i + (k * TimesItBy) + ButtonWidth, 1)-(i + (k * TimesItBy) + ButtonWidth, 13), RGB(RgbColour.R, RgbColour.G, RgbColour.B)
  End If
End Sub
Private Sub DrawR(ByVal i As Long, TimesItBy As Integer)
  If Bar3D = Flat Then
    Line ((BarLength + ButtonWidth) - (i + (k * TimesItBy)), 0)-((BarLength + ButtonWidth) - (i + (k * TimesItBy)), 14), RGB(RgbColour.R, RgbColour.G, RgbColour.B)
  Else
    Line ((BarLength + ButtonWidth) - (i + (k * TimesItBy)), 1)-((BarLength + ButtonWidth) - (i + (k * TimesItBy)), 13), RGB(RgbColour.R, RgbColour.G, RgbColour.B)
  End If
End Sub

Private Function CheckColorNow(Checked As Long) As Boolean
On Error Resume Next
  Dim checkValue As Byte
  checkValue = 230
  CheckedTheColor.R = ((Checked Mod 65536) Mod 256)
  CheckedTheColor.G = ((Checked Mod 65536) \ 256)
  CheckedTheColor.B = (Checked \ 65536)
  
  If CheckedTheColor.R > checkValue Or CheckedTheColor.G > checkValue Then
    CheckColorNow = True
  Else
    CheckColorNow = False
  End If
End Function
Public Sub AutoCalcPercentage(Percent As Integer, ValueToCalc As Integer)
  Dim tempValue As Single
  BarCVMax = ValueToCalc
  tempValue = ValueToCalc / 100
  BarCurrentValue = Int(tempValue * Percent)
  DisplayUp
End Sub
