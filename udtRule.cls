VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "udtRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' m_lngID is the id associated with the rule in the registry
Public m_lngID As Long

' m_strName is the name of the rule
Public m_strName As String

' types
Public m_lngWhenType As Long
Public m_lngWhatType As Long

' parameters
Public m_strWhenParams As String
Public m_strWhatParams As String

