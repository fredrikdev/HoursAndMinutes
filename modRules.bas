Attribute VB_Name = "modRules"
Option Explicit

' about rules:
' - rules are saved as long text strings
' - several rules are separated with * (star)
' - the format for each rule is:
' version|name|when|what|minute-parameter|message-parameter|sound-parameter

' when rule types
Public Const RULE_WHEN_AFTER_X_MINUTES_ONE_DAY = 101
Public Const RULE_WHEN_EVERY_X_MINUTES = 102

' what rule types
Public Const RULE_WHAT_SHOW_MESSAGE = 101
Public Const RULE_WHAT_PLAY_SOUND = 102
Public Const RULE_WHAT_SHOW_MESSAGE_PLAY_SOUND = 103

' system functions
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

' processes rules
Public Sub ruleProcess(lngModeID As Long, lngMSeconds As Long)
    Dim arrRule() As String, arrRules() As String
    Dim arrRulesLastRun() As String
    Dim x As Long
    
    With m_colModes.Item(CStr(lngModeID))
        arrRules = Split(.m_strRule, "*")
        If UBound(arrRules) < 0 Then Exit Sub
        If .m_strRulesLastRun = "" Then
            .m_strRulesLastRun = Repeat("|", UBound(arrRules))
        End If
        arrRulesLastRun = Split(.m_strRulesLastRun, "|")
    
        For x = LBound(arrRules) To UBound(arrRules)
            arrRule = Split(arrRules(x), "|")
            If UBound(arrRule) < 6 Then
            Else
                ' process a version 1.0 rule
                If arrRule(0) = "1.0" Then
                    Select Case CLng(arrRule(2))
                        Case RULE_WHEN_AFTER_X_MINUTES_ONE_DAY
                            If arrRulesLastRun(x) = "" Then
                                If (lngMSeconds > CLng(arrRule(4)) * 60 * 1000) Then
                                    arrRulesLastRun(x) = CLng(Date)
                                    ruleExecute arrRules(x)
                                End If
                            Else
                                If (CDate(arrRulesLastRun(x)) <> Date) And (lngMSeconds > CLng(arrRule(4)) * 60 * 1000) Then
                                    arrRulesLastRun(x) = Date
                                    ruleExecute arrRules(x)
                                End If
                            End If
                        Case RULE_WHEN_EVERY_X_MINUTES
                            If arrRulesLastRun(x) = "" Then arrRulesLastRun(x) = lngMSeconds
                            If (lngMSeconds - CLng(arrRulesLastRun(x))) > CLng(arrRule(4)) * 60 * 1000 Then
                                arrRulesLastRun(x) = lngMSeconds
                                ruleExecute arrRules(x)
                            End If
                    End Select
                End If
            End If
        Next
        
        .m_strRulesLastRun = Join(arrRulesLastRun, "|")
    End With
End Sub

' executes the actions of a rule
Public Sub ruleExecute(strRule As String)
    Dim arrRule() As String
    
    If ruleValidate(strRule) Then
        arrRule = Split(strRule, "|")
        
        ' execute a version 1.0 rule
        If arrRule(0) = "1.0" Then
            Select Case CLng(arrRule(3))
                Case RULE_WHAT_SHOW_MESSAGE
                    showMessage arrRule(5), arrRule(1)
                Case RULE_WHAT_PLAY_SOUND
                    sndPlaySound arrRule(6), &H1
                Case RULE_WHAT_SHOW_MESSAGE_PLAY_SOUND
                    sndPlaySound arrRule(6), &H1
                    showMessage arrRule(5), arrRule(1)
            End Select
        End If
    End If
End Sub

' validates a rule and verifies that all values has been set
Public Function ruleValidate(strRule As String) As Boolean
    Dim arrRule() As String
    
    ruleValidate = False
    arrRule = Split(strRule, "|")
    
    If UBound(arrRule) < 6 Then Exit Function
    
    ' validate a version 1.0 rule
    If arrRule(0) = "1.0" Then
        ' validate when part
        Select Case CLng(arrRule(2))
            Case RULE_WHEN_AFTER_X_MINUTES_ONE_DAY
                If Not isInteger(arrRule(4)) Then Exit Function
            Case RULE_WHEN_EVERY_X_MINUTES
                If Not isInteger(arrRule(4)) Then Exit Function
            Case Else
                Exit Function
        End Select
        
        ' validate what part
        Select Case CLng(arrRule(3))
            Case RULE_WHAT_SHOW_MESSAGE
                If arrRule(5) = "undefined" Or arrRule(5) = "" Then Exit Function
            Case RULE_WHAT_PLAY_SOUND
                If arrRule(6) = "undefined" Or arrRule(6) = "" Then Exit Function
            Case RULE_WHAT_SHOW_MESSAGE_PLAY_SOUND
                If arrRule(5) = "undefined" Or arrRule(5) = "" Then Exit Function
                If arrRule(6) = "undefined" Or arrRule(6) = "" Then Exit Function
            Case Else
                Exit Function
        End Select
        
        ruleValidate = True
        Exit Function
    End If
End Function

' returns the (userdefined)name of a rule
' note: this function requires that the rule has been validate with
'       ruleValidate, since no validation will be performed
Public Function ruleGetName(strRule As String) As String
    ruleGetName = Mid(strRule, InStr(1, strRule, "|") + 1, (InStr(InStr(1, strRule, "|") + 1, strRule, "|") - 1) - InStr(1, strRule, "|"))
End Function

' return the data of a rule - this excludes the version and name
' from the rule.
' note: this function requires that the rule has been validate with
'       ruleValidate, since no validation will be performed
Public Function ruleGetData(strRule As String) As String
    Dim x As Long
    x = InStr(1, strRule, "|") + 1
    x = InStr(x, strRule, "|") + 1
    ruleGetData = Mid(strRule, x)
End Function
