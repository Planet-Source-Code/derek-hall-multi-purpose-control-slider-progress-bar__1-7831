VERSION 5.00
Object = "*\AProject1.vbp"
Begin VB.Form ProgressDemo 
   Caption         =   "Form1"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Calc"
      Height          =   255
      Left            =   3300
      TabIndex        =   31
      Top             =   1620
      Width           =   615
   End
   Begin VB.TextBox txtValue 
      Alignment       =   2  'Center
      Height          =   225
      Left            =   2700
      TabIndex        =   29
      Text            =   "300"
      Top             =   1620
      Width           =   435
   End
   Begin VB.TextBox txtPercentage 
      Alignment       =   2  'Center
      Height          =   225
      Left            =   1680
      TabIndex        =   28
      Text            =   "44"
      Top             =   1620
      Width           =   495
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   6720
      Top             =   1800
   End
   Begin VB.TextBox CValue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   5580
      TabIndex        =   23
      Text            =   "0"
      Top             =   1980
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6240
      Top             =   1800
   End
   Begin VB.ComboBox cmbShowValue 
      Height          =   315
      ItemData        =   "ProgressDemo.frx":0000
      Left            =   7380
      List            =   "ProgressDemo.frx":000A
      TabIndex        =   20
      Top             =   360
      Width           =   1035
   End
   Begin VB.TextBox CValue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   3300
      TabIndex        =   14
      Text            =   "0"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox CValue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   3300
      TabIndex        =   13
      Text            =   "0"
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox CValue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   3300
      TabIndex        =   12
      Text            =   "0"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox CValue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   3300
      TabIndex        =   11
      Text            =   "0"
      Top             =   1380
      Width           =   615
   End
   Begin VB.ComboBox cmbColor 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "ProgressDemo.frx":0017
      Left            =   4200
      List            =   "ProgressDemo.frx":0033
      TabIndex        =   10
      Text            =   "Select A Color"
      Top             =   360
      Width           =   1155
   End
   Begin VB.ComboBox cmbStyle 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "ProgressDemo.frx":007F
      Left            =   6240
      List            =   "ProgressDemo.frx":008C
      TabIndex        =   9
      Text            =   "Select Border Style"
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox txtCustomColour 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   8
      Text            =   "0"
      Top             =   1140
      Width           =   1935
   End
   Begin VB.CommandButton cmdRND 
      Caption         =   "RND"
      Height          =   315
      Left            =   5400
      TabIndex        =   7
      Top             =   360
      Width           =   735
   End
   Begin VB.Frame optBar 
      Height          =   1995
      Left            =   0
      TabIndex        =   2
      Top             =   420
      Width           =   435
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   195
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Value           =   -1  'True
         Width           =   195
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   195
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   660
         Width           =   195
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   195
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4200
      Picture         =   "ProgressDemo.frx":00A4
      ScaleHeight     =   240
      ScaleWidth      =   1920
      TabIndex        =   1
      Top             =   780
      Width           =   1920
   End
   Begin VB.ComboBox cmb_Direction 
      Height          =   315
      ItemData        =   "ProgressDemo.frx":18E6
      Left            =   6240
      List            =   "ProgressDemo.frx":18F0
      TabIndex        =   0
      Top             =   1020
      Width           =   1095
   End
   Begin Project1.ProgressBar ProgressBar1 
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   15
      Top             =   600
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   238
      ColourScheme    =   0
      BarValue        =   100
      BarMaxValue     =   255
      Bar3D           =   1
      BarBChange      =   10
      BarPreValue     =   255
      ButtonsAreOn    =   1
      BarBackColor    =   8421504
      BarControl      =   0
   End
   Begin Project1.ProgressBar ProgressBar1 
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   24
      Top             =   840
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   238
      BarValue        =   20
      BarMaxValue     =   20
      Bar3D           =   1
      BarBChange      =   10
      BarPreValue     =   100
      ButtonsAreOn    =   1
      BarBackColor    =   11837122
   End
   Begin Project1.ProgressBar ProgressBar1 
      Height          =   195
      Index           =   2
      Left            =   600
      TabIndex        =   25
      Top             =   1080
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   238
      ColourScheme    =   4
      BarValue        =   100
      BarMaxValue     =   100
      Bar3D           =   1
      BarBChange      =   10
      BarPreValue     =   255
      ButtonsAreOn    =   0
      BarBackColor    =   255
      BarControl      =   2
   End
   Begin Project1.ProgressBar ProgressBar1 
      Height          =   195
      Index           =   3
      Left            =   600
      TabIndex        =   26
      Top             =   1380
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   238
      ColourScheme    =   5
      BarMaxValue     =   0
      Bar3D           =   1
      BarBChange      =   10
      ButtonsAreOn    =   0
      BarBackColor    =   255
      BarControl      =   0
   End
   Begin Project1.ProgressBar ProgressBar1 
      Height          =   195
      Index           =   4
      Left            =   600
      TabIndex        =   27
      Top             =   1980
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   344
      ColourScheme    =   6
      BarValue        =   100
      BarMaxValue     =   100
      Bar3D           =   1
      BarBChange      =   10
      BarPreValue     =   255
      ButtonsAreOn    =   0
      BarBackColor    =   255
      BarControl      =   2
   End
   Begin VB.Shape Shape2 
      Height          =   615
      Left            =   540
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label Label7 
      Caption         =   "Of"
      Height          =   255
      Left            =   2340
      TabIndex        =   32
      Top             =   1620
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "Percentage"
      Height          =   255
      Left            =   780
      TabIndex        =   30
      Top             =   1620
      Width           =   915
   End
   Begin VB.Label Label5 
      Caption         =   "Show Value"
      Height          =   255
      Left            =   7380
      TabIndex        =   21
      Top             =   120
      Width           =   1035
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   4140
      Top             =   300
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Style"
      Height          =   195
      Left            =   4140
      TabIndex        =   19
      Top             =   60
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Appearance"
      Height          =   195
      Left            =   6300
      TabIndex        =   18
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Previous Value"
      Height          =   375
      Left            =   3300
      TabIndex        =   17
      Top             =   180
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "Reverse Shade"
      Height          =   255
      Left            =   6240
      TabIndex        =   16
      Top             =   780
      Width           =   1215
   End
End
Attribute VB_Name = "ProgressDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  Dim CurrentBar As Integer

Private Sub cmb_Direction_Click()
  ProgressBar1(CurrentBar).BarReverseShade = cmb_Direction.ItemData(cmb_Direction.ListIndex)
End Sub

Private Sub cmbColor_Click()
  ProgressBar1(CurrentBar).BarStyle = cmbColor.ItemData(cmbColor.ListIndex)
  UpdateControls
End Sub

Private Sub cmbShowValue_Click()
  ProgressBar1(CurrentBar).BarShowValue = cmbShowValue.ItemData(cmbShowValue.ListIndex)
End Sub


Private Sub cmbStyle_Click()
  ProgressBar1(CurrentBar).BarAppearance = cmbStyle.ItemData(cmbStyle.ListIndex)
End Sub

Private Sub cmdRND_Click()
  txtCustomColour = 16777216 * Rnd
End Sub


Private Sub Command2_Click()
Form1.Show
End Sub

Private Sub Command1_Click()
ProgressBar1(3).AutoCalcPercentage txtPercentage, txtValue
End Sub

Private Sub Form_Load()
CurrentBar = 0
  Dim i As Integer
  For i = 3 To 0 Step -1
    ProgressBar1_Click i
    UpdateControls
  Next i
End Sub
Sub UpdateControls()
  cmbColor.ListIndex = ProgressBar1(CurrentBar).BarStyle
  txtCustomColour = ProgressBar1(CurrentBar).BarCustomColor
  cmbStyle.ListIndex = ProgressBar1(CurrentBar).BarAppearance
  cmb_Direction.ListIndex = ProgressBar1(CurrentBar).BarReverseShade
  
  cmbShowValue.ListIndex = ProgressBar1(CurrentBar).BarShowValue
  
  If cmbColor = "Custom" Or cmbColor = "Standard" Or cmbColor = "DialType" Then
    txtCustomColour.Visible = True
    cmdRND.Visible = True
    Picture1.Visible = True
  Else
    txtCustomColour.Visible = False
    cmdRND.Visible = False
    Picture1.Visible = False
  End If
End Sub


Private Sub Option1_Click(index As Integer)
  CurrentBar = index
  UpdateControls
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  txtCustomColour = Picture1.Point(X, Y)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then txtCustomColour = Picture1.Point(X, Y)
End Sub

Private Sub ProgressBar1_Click(index As Integer)
  CValue(index) = ProgressBar1(index).BarPreviousValue
 ' Option1(index) = True
  'UpdateControls
End Sub

Private Sub Timer1_Timer()
ProgressBar1(2).BarCurrentValue = ProgressBar1(2).BarCurrentValue + 1
If ProgressBar1(2).BarCurrentValue = 100 Then ProgressBar1(2).BarCurrentValue = 0
ProgressBar1(4).BarCurrentValue = ProgressBar1(4).BarCurrentValue - 1
End Sub

Private Sub Timer2_Timer()
ProgressBar1(4).BarCurrentValue = Int(Rnd * 75) + 25
End Sub

Private Sub txtCustomColour_Change()
  On Error Resume Next ' quick way to handle error if not a Long
  ProgressBar1(CurrentBar).BarCustomColor = txtCustomColour
End Sub

