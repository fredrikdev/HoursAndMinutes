Attribute VB_Name = "modModes"
Option Explicit

' m_colModes is a collection of udtMode classes
Public m_colModes As Collection

' this initializes the mode handler
Public Sub modesInitialize()
    Dim objReg As clsRegistry
    Dim arrKeys() As String, lngCount As Long, X As Long
    Dim strName As String, strRule As String
    Dim objMode As udtMode
    
    Set m_colModes = New Collection
    
    ' load modes & rules from the registry
    Set objReg = New clsRegistry
    With objReg
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = m_strRegRoot & "\Modes"
        
        If .EnumerateSections(arrKeys, lngCount) Then
            For X = 1 To lngCount
                .SectionKey = m_strRegRoot & "\Modes\" & arrKeys(X)
                
                .ValueKey = "m_strName"
                strName = CStr(.Value)
                
                .ValueKey = "m_strRule"
                strRule = CStr(.Value)
                                
                Set objMode = New udtMode
                objMode.m_lngID = CLng(arrKeys(X))
                objMode.m_strName = strName
                objMode.m_strRule = strRule
                m_colModes.Add objMode, CStr(arrKeys(X))
                Set objMode = Nothing
            Next
        End If
    End With
    Set objReg = Nothing
End Sub

' this stores the modes collection to the registry
Public Sub modesStore()
    Dim objReg As clsRegistry
    Dim arrKeys() As String, lngCount As Long, X As Long
    
    Set objReg = New clsRegistry
    With objReg
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = m_strRegRoot & "\Modes"
        
        ' get all sections that exists in the registry
        If .EnumerateSections(arrKeys, lngCount) Then
            ' remove sections from the registry that does not exists in our collection
            For X = 1 To lngCount
                If isInCollection(arrKeys(X), m_colModes) = False Then
                    .SectionKey = m_strRegRoot & "\Modes\" & arrKeys(X)
                    .DeleteKeyRec
                End If
            Next
        End If
        
        ' update/add sections
        For X = 1 To m_colModes.Count
            With m_colModes.Item(X)
                objReg.SectionKey = m_strRegRoot & "\Modes\" & .m_lngID
                objReg.ValueKey = "m_strName"
                objReg.Value = .m_strName
                objReg.ValueKey = "m_strRule"
                objReg.Value = .m_strRule
            End With
        Next
    End With
    Set objReg = Nothing
End Sub

' this adds/updates a new mode to the collection
Public Sub modeUpdate(strName As String, strRule As String, Optional lngID As Long = -1)
    Dim objMode As udtMode
    
    If lngID = -1 Then
        Set objMode = New udtMode
        With objMode
            .m_lngID = modeGetNextID
            .m_strName = strName
            .m_strRule = strRule
            .m_strRulesLastRun = ""
        End With
        m_colModes.Add objMode, CStr(objMode.m_lngID)
        Set objMode = Nothing
    Else
        With m_colModes.Item(CStr(lngID))
            If .m_strRule <> strRule Then
                .m_strRulesLastRun = ""
            End If
            .m_strName = strName
            .m_strRule = strRule
        End With
    End If
End Sub

' this saves mode time into the registry
Public Sub modeSaveTime(lngID As Long, dteDate As Date, lngTime As Long)
    Dim objReg As clsRegistry
    
    If lngTime < 0 Then Exit Sub
    If frmMain.m_blnLoadState Then Exit Sub
    
    Set objReg = New clsRegistry
    With objReg
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = m_strRegRoot & "\Modes\" & lngID
        .ValueKey = "Date " & CLng(dteDate)
        .Value = lngTime
    End With
    Set objReg = Nothing
End Sub

' this get mode time from the registry
Public Function modeGetTime(lngID As Long, dteDate As Date) As Long
    Dim objReg As clsRegistry
    
    Set objReg = New clsRegistry
    With objReg
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = m_strRegRoot & "\Modes\" & lngID
        .ValueKey = "Date " & CLng(dteDate)
        modeGetTime = .Value
    End With
    Set objReg = Nothing
    
End Function

' this removes a mode from the collection
Public Sub modeDelete(lngID As Long)
    m_colModes.Remove CStr(lngID)
End Sub

' this retrieves the next available mode id, and updates the registry
Private Function modeGetNextID() As Long
    Dim objReg As clsRegistry
    Dim lngTemp As Long
    
    Set objReg = New clsRegistry
    With objReg
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = m_strRegRoot & "\Modes"
        
        .ValueKey = "m_lngNextModeID"
        lngTemp = .Value
        
        modeGetNextID = lngTemp
        
        lngTemp = lngTemp + 1
        .Value = lngTemp
    End With
    Set objReg = Nothing
End Function

