VERSION 5.00
Begin VB.UserControl ctlTimer 
   BackColor       =   &H00FF8080&
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   420
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
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   420
   ScaleWidth      =   420
   Begin VB.Timer ctlTimer 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "ctlTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' private declares:
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Type POINTAPI
    x As Long
    y As Long
End Type

' private variables:
Private m_dteDate As Date
Private m_lngTStart As Long
Private m_blnIdle As Boolean
Private m_lngTLastAutosave As Long
' public variables:
Public m_blnAllowIdle As Boolean
Public m_lngIdleStart As Long
Public m_lngIdleEnd As Long

Public m_blnAllowAutosave As Boolean
Public m_lngAutosaveInterval As Long
' public events
Public Event Timer(dteDate As Date, dteNextDate As Date, lngMSeconds As Long, blnIdle As Boolean, blnAutosaveNow As Boolean, blnFinalResults As Boolean)

Public Sub Start(Optional lngTime As Long = 0)
    m_dteDate = Date
    m_lngTStart = GetTickCount - lngTime
    m_lngTLastAutosave = GetTickCount
End Sub

Private Sub ctlTimer_Timer()
    Dim blnDateWillChange As Boolean, lngITime As Long, lngTCount As Long
    Dim blnAutosave As Boolean
    Static m_lngResumed As Variant, m_lngIdleTime As Long
    
    lngTCount = GetTickCount
    lngITime = getIdleTime(lngTCount)
    
    blnAutosave = False
    If m_blnAllowAutosave Then
        If (lngTCount - m_lngTLastAutosave) \ 1000 \ 60 >= m_lngAutosaveInterval Then
            blnAutosave = True
            m_lngTLastAutosave = lngTCount
        End If
    End If
    
    ' check if idle
    If m_blnAllowIdle = True Then
        If (lngITime >= (m_lngIdleStart * 60)) Then
            If m_blnIdle = False Then
                m_blnIdle = True
                m_lngResumed = Empty
                m_lngIdleTime = lngTCount
            End If
        Else
            If m_blnIdle = True Then
                If IsEmpty(m_lngResumed) Then m_lngResumed = lngTCount
                If ((lngTCount - m_lngResumed) \ 1000) >= (m_lngIdleEnd * 60) Then
                    m_blnIdle = False
                    m_lngTStart = m_lngTStart + (lngTCount - m_lngIdleTime)
                End If
            End If
        End If
    Else
        m_blnIdle = False
    End If
    
    ' fire event
    blnDateWillChange = m_dteDate <> Date
    If m_blnIdle Then
        RaiseEvent Timer(m_dteDate, Date, (lngTCount - m_lngTStart) - (lngTCount - m_lngIdleTime), m_blnIdle, blnAutosave, blnDateWillChange)
    Else
        RaiseEvent Timer(m_dteDate, Date, lngTCount - m_lngTStart, m_blnIdle, blnAutosave, blnDateWillChange)
    End If
    
    ' reset counter
    If blnDateWillChange Then
'        m_dteDate = Date
'        m_lngTStart = lngTCount
        If m_blnIdle Then
            m_lngIdleTime = lngTCount
        Else
            m_lngIdleTime = 0
        End If
    End If
End Sub

' returns the idle time in seconds (if called once every second)
Private Function getIdleTime(lngTCount As Long) As Long
    Static objOldPoint As POINTAPI
    Static objLastRun As Variant
    Dim objPoint As POINTAPI, x As Integer, blnTemp As Boolean
    getIdleTime = 0
    
    ' check if cursor has been moved since last run
    GetCursorPos objPoint
    
    blnTemp = False
    If objPoint.x <> objOldPoint.x Or objPoint.y <> objOldPoint.y Then
        blnTemp = True
    End If
    For x = 1 To 256
        If GetAsyncKeyState(x) <> 0 Then
            blnTemp = True
        End If
    Next
    
    If blnTemp = False Then
        ' idle
        If IsEmpty(objLastRun) Then objLastRun = lngTCount
        getIdleTime = ((lngTCount - objLastRun) \ 1000)
    Else
        ' active
        objOldPoint = objPoint
        objLastRun = Empty
    End If
End Function

