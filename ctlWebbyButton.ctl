VERSION 5.00
Begin VB.UserControl ctlWebbyButton 
   BackColor       =   &H00AC7C67&
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1980
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   MouseIcon       =   "ctlWebbyButton.ctx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   21
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   132
   Begin VB.Image imgArrow 
      Height          =   90
      Left            =   165
      Picture         =   "ctlWebbyButton.ctx":030A
      Top             =   120
      Width           =   120
      Visible         =   0   'False
   End
End
Attribute VB_Name = "ctlWebbyButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

'Default Property Values:
Const m_def_Selected = False
Const m_def_Caption = "Webby"
Dim m_colBackground As Long, m_colBackgroundHilight As Long, m_colBorder As Long
'Property Variables:
Dim m_Selected As Boolean
Dim m_Caption As String
Public Event Click()

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_Initialize()
    ' setup colors
    m_colBackground = RGB(103, 124, 172)
    m_colBackgroundHilight = RGB(255, 255, 255)
    m_colBorder = RGB(70, 86, 120)
    
    UserControl.BackColor = m_colBackground
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Selected Then
        If UserControl.BackColor <> m_colBackground Then
            UserControl.BackColor = m_colBackground
            UserControl.ForeColor = m_colBackgroundHilight
            UserControl_Paint
        End If
        Exit Sub
    End If
    If (x < 0) Or (y < 0) Or (x > UserControl.ScaleWidth) Or (y > UserControl.ScaleHeight) Or (Button <> 0) Then
        ReleaseCapture
        If UserControl.BackColor <> m_colBackground Then
            UserControl.BackColor = m_colBackground
            UserControl.ForeColor = m_colBackgroundHilight
            UserControl_Paint
        End If
    Else
        SetCapture UserControl.hwnd
        If UserControl.BackColor <> m_colBackgroundHilight Then
            UserControl.BackColor = m_colBackgroundHilight
            UserControl.ForeColor = m_colBackground
            UserControl_Paint
        End If
    End If
End Sub

Private Sub UserControl_Paint()
    UserControl.Cls
    UserControl.Line (0, 0)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), m_colBorder, B
    UserControl.CurrentX = 25
    UserControl.CurrentY = 4
    UserControl.Print m_Caption
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 132 * Screen.TwipsPerPixelX
    UserControl.Height = 21 * Screen.TwipsPerPixelY
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,"Webby"
Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    PropertyChanged "Caption"
    UserControl_Paint
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Caption = m_def_Caption
    m_Selected = m_def_Selected
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    m_Selected = PropBag.ReadProperty("Selected", m_def_Selected)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Selected", m_Selected, m_def_Selected)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Selected() As Boolean
    Selected = m_Selected
End Property

Public Property Let Selected(ByVal New_Selected As Boolean)
    m_Selected = New_Selected
    PropertyChanged "Selected"
    imgArrow.Visible = New_Selected
End Property

