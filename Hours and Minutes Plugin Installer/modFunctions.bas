Attribute VB_Name = "modFunctions"
Option Explicit

Private Declare Function DLLSelfRegister Lib "VB6STKIT.DLL" (ByVal lpDllName As String) As Integer

Public Sub Main()
    Dim arrParameters() As String
    Dim strFilePath As String, strFileName As String, strCreateString As String
    Dim strMyPath As String
    Dim objPlugin As Object, objReg As clsRegistry
    Dim strPCaption As String
    Dim x As Long, y As Long, z As Long
    Dim arrValues() As String
    
    strMyPath = App.Path
    If Right(strMyPath, 1) <> "\" Then strMyPath = strMyPath & "\"
    strMyPath = strMyPath & "Plugins\"
    On Error Resume Next
    MkDir strMyPath
    On Error GoTo 0
    
    On Error Resume Next
    arrParameters = ParseParameters
    If UBound(arrParameters) = -1 Then
        ShowProgress "Error: Failed to parse parameters!"
        Exit Sub
    End If
    For x = LBound(arrParameters) To UBound(arrParameters)
        strFilePath = ExpandPath(arrParameters(x))
        strFileName = Mid$(strFilePath, InStrRev(strFilePath, "\") + 1)
        strFilePath = Left(strFilePath, InStrRev(strFilePath, "\"))
        strCreateString = Left(strFileName, InStrRev(strFileName, ".") - 1)
        If FileExists(strFilePath & strFileName) Then
            If LCase(strFilePath) = LCase(strMyPath) Then
                ' uninstall plugin
                On Error GoTo lblErrorUnInstall
                
                ' check parameters
                If Trim(strFilePath) = "" Then Err.Raise 2005
                If Trim(strFileName) = "" Then Err.Raise 2006
                If Trim(strCreateString) = "" Then Err.Raise 2007

                ' remove registration entries
                On Error Resume Next
                Set objReg = New clsRegistry
                With objReg
                    .ClassKey = HKEY_CURRENT_USER
                    .SectionKey = "Software\Port Jackson Computing\Hours and Minutes\Plugins\Statistics"
                    If (.EnumerateValues(arrValues, y) = True) Then
                        For z = 1 To y
                            .ValueKey = arrValues(z)
                            If LCase(.Value) = LCase(strCreateString) Then
                                .DeleteValue
                                Exit For
                            End If
                        Next
                    End If
                End With
                Set objReg = Nothing
                
                ' query plugin
                Set objPlugin = CreateObject(strCreateString)
                If objPlugin.PluginType <> "Hours and Minutes Statistics Plugin 1.0" Then Err.Raise 2001
                strPCaption = objPlugin.PluginCaption
                Set objPlugin = Nothing
                                
                On Error GoTo lblErrorUnInstall
                
                ' remove the plugin
                Kill strFilePath & strFileName
                
                ShowProgress "The Plugin '" & strPCaption & "' has been successfully removed!"
                                
                GoTo lblNext
lblErrorUnInstall:
                Select Case Err.Number
                    Case 2001: ShowProgress "Error 1001: This plugin was created for a newer version of Hours and Minutes"
                    Case 2002: ShowProgress "Error 1002: This plugin was created for a newer version of Hours and Minutes"
                    Case 2005: ShowProgress "Error 1005: Invalid Source Path"
                    Case 2006: ShowProgress "Error 1006: Invalid Filename"
                    Case 2007: ShowProgress "Error 1007: Invalid Filename (name must be <project>.<class>.phm)"
                    Case Else: ShowProgress "Error: " & Err.Description
                End Select
                Exit Sub
            Else
                On Error GoTo lblErrorInstall:
                ' check parameters
                If Trim(strFilePath) = "" Then Err.Raise 1005
                If Trim(strFileName) = "" Then Err.Raise 1006
                If Trim(strCreateString) = "" Then Err.Raise 1007
            
                ' copy plugin
                FileCopy strFilePath & strFileName, strMyPath & strFileName
                
                ' register plugin
                If DLLSelfRegister(strMyPath & strFileName) <> 0 Then Err.Raise 1000
                
                ' query plugin
                Set objPlugin = CreateObject(strCreateString)
                If objPlugin.PluginType <> "Hours and Minutes Statistics Plugin 1.0" Then Err.Raise 1001
                strPCaption = objPlugin.PluginCaption
                If Trim(strPCaption) = "" Then Err.Raise 1003
                Set objPlugin = Nothing
                
                ' register within Hours and Minutes
                Set objReg = New clsRegistry
                With objReg
                    .ClassKey = HKEY_CURRENT_USER
                    .SectionKey = "Software\Port Jackson Computing\Hours and Minutes\Plugins\Statistics"
                    .ValueKey = strPCaption
                    .Value = strCreateString
                End With
                Set objReg = Nothing
                                
                ShowProgress "The Plugin '" & strPCaption & "' has been successfully installed!" & vbCrLf & "You may now remove the source file."
                GoTo lblNext
lblErrorInstall:
                Select Case Err.Number
                    Case 1000: ShowProgress "Error 1000: Failed while registering DLL"
                    Case 1001: ShowProgress "Error 1001: This plugin was created for a newer version of Hours and Minutes"
                    Case 1002: ShowProgress "Error 1002: This plugin was created for a newer version of Hours and Minutes"
                    Case 1003: ShowProgress "Error 1003: This plugin has no caption or was created for a newer version of Hours and Minutes"
                    Case 1005: ShowProgress "Error 1005: Invalid Source Path"
                    Case 1006: ShowProgress "Error 1006: Invalid Filename"
                    Case 1007: ShowProgress "Error 1007: Invalid Filename (name must be <project>.<class>.phm)"
                    Case Else: ShowProgress "Error: " & Err.Description
                End Select
                Exit Sub
            End If
lblNext:
        End If
    Next
End Sub

Public Function ParseParameters() As String()
    Dim blnInSwitch As Boolean, strCommands As String, x As Long
    Dim arrCommands() As String
    
    ' parse commands
    blnInSwitch = False
    strCommands = ""
    If InStr(1, Command$, """") < 1 Then
        ParseParameters = Split(Command$, "|")
        Exit Function
    End If
    For x = 1 To Len(Command$)
        If Mid$(Command$, x, 1) = """" Then
            If blnInSwitch Then
                strCommands = strCommands & "|"
            End If
            blnInSwitch = Not blnInSwitch
        ElseIf blnInSwitch Then
            strCommands = strCommands & Mid$(Command$, x, 1)
        End If
    Next
    If Right(strCommands, 1) = "|" Then strCommands = Left(strCommands, Len(strCommands) - 1)
    
    ' trim commands
    arrCommands = Split(strCommands, "|")
    For x = LBound(arrCommands) To UBound(arrCommands)
        arrCommands(x) = Trim(arrCommands(x))
    Next
    
    ParseParameters = arrCommands
End Function

Public Sub ShowProgress(strText As String)
    frmMessage.lblCaption.Caption = strText
    frmMessage.Show 1
End Sub

Public Function FileExists(strFile As String) As Boolean
    On Error GoTo lblError
    FileExists = False
    If FileLen(strFile) > 0 Then FileExists = True
    Exit Function
lblError:
    Err.Clear
    On Error GoTo 0
End Function

Public Function ExpandPath(strPath As String) As String
    Dim arrPath() As String, strRes As String
    Dim x As Long
    
    ExpandPath = strPath
    arrPath = Split(strPath, "\")
    If UBound(arrPath) = 0 Then Exit Function
    strRes = arrPath(0) & "\"
    For x = LBound(arrPath) + 1 To UBound(arrPath)
        strRes = strRes & Dir(strRes & arrPath(x), vbDirectory Or vbHidden Or vbNormal Or vbReadOnly Or vbSystem) & "\"
    Next
    If Right(strRes, 1) = "\" Then strRes = Left(strRes, Len(strRes) - 1)
    
    ExpandPath = strRes
End Function
