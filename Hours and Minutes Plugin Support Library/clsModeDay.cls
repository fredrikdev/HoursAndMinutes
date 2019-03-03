VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsModeDay"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_dteDate As Date
Private m_lngMSeconds As Long
Private m_arrComments() As String

Public Property Get DDate() As Date
    DDate = m_dteDate
End Property

Public Property Let DDate(dteValue As Date)
    m_dteDate = dteValue
End Property

Public Property Get Milliseconds() As Long
    Milliseconds = m_lngMSeconds
End Property

Public Property Let Milliseconds(lngValue As Long)
    m_lngMSeconds = lngValue
End Property

Public Property Get Comments() As String()
    Comments = m_arrComments
End Property

Public Property Let Comments(strValue() As String)
    m_arrComments = strValue
End Property

