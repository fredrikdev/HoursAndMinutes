VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMode"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_strModeName As String
Private m_arrDays() As clsModeDay

Public Property Get modeName() As String
    modeName = m_strModeName
End Property

Public Property Let modeName(strValue As String)
    m_strModeName = strValue
End Property

Public Property Get modeDays() As clsModeDay()
    modeDays = m_arrDays
End Property

Public Property Let modeDays(arrValue() As clsModeDay)
    m_arrDays = arrValue
End Property

