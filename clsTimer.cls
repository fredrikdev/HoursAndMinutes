VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' private variables
Private m_sngTimerStart As Single
Private m_sngStartValue As Single
Private m_dteDate As Date

' public events
Public Event DayChanged(sngYesterdayTime As Single)

' function start - starts the timing operation
' adds milliseconds to the number of milliseconds in sngStartValue
Public Sub Start(Optional sngStartValue As Single = 0)
    m_sngStartValue = sngStartValue
    m_sngTimerStart = Timer
End Sub

' function gettime - returns time elapsed
' the return value is in milliseconds
Public Function GetTime() As Single
    Dim sngTemp As Single
End Function
