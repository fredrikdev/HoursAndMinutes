VERSION 5.00
Begin VB.Form frmModeMessage 
   BorderStyle     =   0  'None
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2220
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "frmModeMessage.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   120
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   148
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrHide 
      Interval        =   5000
      Left            =   510
      Top             =   1230
   End
   Begin VB.Timer tmrMessage 
      Left            =   930
      Top             =   1230
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "The selected mode is Working with This.                                           Click here to switch mode."
      ForeColor       =   &H00FFFFFF&
      Height          =   1725
      Left            =   510
      MouseIcon       =   "frmModeMessage.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   30
      UseMnemonic     =   0   'False
      Width           =   1665
   End
End
Attribute VB_Name = "frmModeMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As RECT, ByVal fuWinIni As Long) As Long
Private Declare Function AppBarMessage Lib "shell32.dll" Alias "SHAppBarMessage" (ByVal dwMessage As Long, pData As APPBARDATA) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type APPBARDATA
    cbSize As Long
    hWnd As Long
    uCallbackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long '  message specific
End Type

Const m_conHeight = 1800
Dim m_sngMessageTimer As Single
Dim m_lngExeCnt As Long
Dim m_blnExpandDown As Boolean
Dim m_blnShow As Boolean
Dim m_sngTop As Single

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (X < 0) Or (X >= ScaleWidth) Or (Y < 0) Or (Y >= ScaleHeight) Then
        ReleaseCapture
        If lblMessage.Font.Underline Then lblMessage.Font.Underline = False
    Else
        SetCapture Me.hWnd
        If Not lblMessage.Font.Underline Then lblMessage.Font.Underline = True
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = MouseButtonConstants.vbLeftButton Then
        frmMain.ctlSwitchModeTimer.Enabled = True
        tmrHide.Enabled = False
        Unload Me
    Else
        tmrMessage.Enabled = False
        tmrHide.Enabled = False
        Unload Me
    End If
End Sub

Private Sub lblMessage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Form_MouseMove(Button, Shift, lblMessage.Left + X / Screen.TwipsPerPixelX, lblMessage.Top + Y / Screen.TwipsPerPixelY)
End Sub

Public Sub showMessage(strMessage As String)
    Dim udtBarData As APPBARDATA
    Dim lngState As Long
    Const ABM_GETTASKBARPOS = 5, ABM_GETSTATE = 4
    Const ABS_ALWAYSONTOP = &H2, ABS_AUTOHIDE = &H1, ABS_BOTH = &H3
    Dim strPath As String
    
    tmrHide.Enabled = False
    tmrMessage.Enabled = False
    
    ' load background
    strPath = App.Path
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    strPath = strPath & "Hours and Minutes Image Switch Reminder.gif"
    Me.Picture = LoadPicture(strPath)
    
    ' get taskbar position
    udtBarData.cbSize = Len(udtBarData)
    lngState = AppBarMessage(ABM_GETSTATE, udtBarData)
    
    AppBarMessage ABM_GETTASKBARPOS, udtBarData
    With udtBarData
        .rc.Bottom = .rc.Bottom * Screen.TwipsPerPixelY
        .rc.Top = .rc.Top * Screen.TwipsPerPixelY
        .rc.Left = .rc.Left * Screen.TwipsPerPixelX
        .rc.Right = .rc.Right * Screen.TwipsPerPixelX
        
        If .uEdge = 0 Then
            ' glued to the left
            Top = Screen.Height
            Left = Screen.Width - Width - Screen.TwipsPerPixelX * 16
            m_blnExpandDown = False
        ElseIf .uEdge = 1 Then
            ' glued to the top
            Top = .rc.Bottom
            Left = Screen.Width - Width - Screen.TwipsPerPixelX * 16
            m_blnExpandDown = True
        ElseIf .uEdge = 2 Then
            ' glued to the right
            Top = Screen.Height
            Left = .rc.Left - Width - Screen.TwipsPerPixelX * 16
            m_blnExpandDown = False
        Else
            ' glued to the bottom
            Top = .rc.Top
            Left = Screen.Width - Width - Screen.TwipsPerPixelX * 16
            m_blnExpandDown = False
        End If
    End With
    m_sngTop = Top
    Height = 0
    m_blnShow = True

    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE
    
    ' set caption
    Set MouseIcon = lblMessage.MouseIcon
    MousePointer = MousePointerConstants.vbCustom
    lblMessage.Caption = strMessage
    lblMessage.Font.Underline = False
    lblMessage.ForeColor = vbWhite
    
    Show
    
    ' initialize
    m_lngExeCnt = 0
    m_sngMessageTimer = Timer
    
    ' start timer
    tmrMessage.Interval = 25 ' 40 executions/sec
    tmrMessage.Enabled = True
End Sub

Private Sub tmrHide_Timer()
    ' disable hide timer
    tmrHide.Enabled = False
    
    ' initialize
    m_blnShow = False
    m_lngExeCnt = 0
    m_sngMessageTimer = Timer
    
    ' hide the form
    tmrMessage.Enabled = True
End Sub

Private Sub tmrMessage_Timer()
    Const lngSpeed = 200
    Dim sngElapsedTime As Single
    Dim lngHeight As Long
    
    m_lngExeCnt = m_lngExeCnt + 1
    sngElapsedTime = Timer - m_sngMessageTimer
    lngHeight = (sngElapsedTime * lngSpeed) * Screen.TwipsPerPixelY
    If m_blnShow Then
        ' show the message
        If lngHeight > m_conHeight Then
            lngHeight = m_conHeight
            tmrMessage.Enabled = False
            tmrHide.Enabled = True
        End If
        
        If m_blnExpandDown Then
            Height = lngHeight
        Else
            Move Left, m_sngTop - lngHeight, Width, lngHeight
        End If
    Else
        ' hide the message
        lngHeight = m_conHeight - lngHeight
        If lngHeight < 0 Then
            lngHeight = 0
            tmrMessage.Enabled = False
            Unload Me
            Exit Sub
        End If
        
        If m_blnExpandDown Then
            Height = lngHeight
        Else
            Move Left, m_sngTop - lngHeight, Width, lngHeight
        End If
    End If
End Sub

