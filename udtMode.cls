VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "udtMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' m_lngID is the id associated with the mode in the registry
Public m_lngID As Long

' m_strName is the name of the mode
Public m_strName As String

' m_strRule is the rule(s) associated with the mode
Public m_strRule As String

' m_strRuleTemp is the rule temporary space, and is NOT saved into the registry
Public m_strRulesLastRun As String

