Attribute VB_Name = "modReminders"
Option Explicit

' about reminders:
' - reminders are saved as long text strings
' - several reminders are separated with * (star)
' - the format for each reminder is:
' version|name|when-mask|when-date|when-time|what-mask|what-message|what-sound|(lastexecuted)

' when reminder types
Public Enum enmReminderWhen
    REMIND_AT_SPECIFIED_DATE_AND_TIME = 1
    REMIND_EVERYDAY_AT_SPECIFIED_TIME = 2
    REMIND_EVERY_WEEKDAY_AT_SPECIFIED_TIME = 4
    REMIND_EVERY_YEAR_AT_SPECIFIED_DATE_AND_TIME = 8
    REMIND_EVERY_MONTH_AT_SPECIFIED_DATE_AND_TIME = 16
End Enum

' what reminder types
Public Enum enmReminderWhat
    SHOW_MESSAGE = 1
    PLAY_SOUND = 2
End Enum

' system functions
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

' process reminders
Public Sub remindersProcess()
    Dim X As Long, arrReminder() As String
    Dim blnTriggerWhat As Boolean, blnRun As Boolean
        
    For X = 1 To frmMain.lstReminders.ListItems.Count
        arrReminder = Split(frmMain.lstReminders.ListItems(X).Tag, "|")
        
        If UBound(arrReminder) = 8 Then
            ' process reminder version 1.0
            If arrReminder(0) = "1.0" Then
                ' determine when...
                blnTriggerWhat = False
                blnRun = False
                If (frmMain.lstReminders.ListItems(X).SubItems(2) = "") Then
                    blnRun = True
                ElseIf (CDate(frmMain.lstReminders.ListItems(X).SubItems(2)) < Date) Then
                    blnRun = True
                End If
                
                If blnRun Then
                    If (CLng(arrReminder(2)) And enmReminderWhen.REMIND_AT_SPECIFIED_DATE_AND_TIME) = enmReminderWhen.REMIND_AT_SPECIFIED_DATE_AND_TIME Then
                        If Date = CDate(arrReminder(3)) Then
                            If Time >= CDate(arrReminder(4)) Then
                                blnTriggerWhat = True
                            End If
                        Else
                            frmMain.lstReminders.ListItems(X).SubItems(2) = Date
                        End If
                    End If
                    If (CLng(arrReminder(2)) And enmReminderWhen.REMIND_EVERY_WEEKDAY_AT_SPECIFIED_TIME) = enmReminderWhen.REMIND_EVERY_WEEKDAY_AT_SPECIFIED_TIME Then
                        If Weekday(Date) = Weekday(CDate(arrReminder(3))) Then
                            If Time >= CDate(arrReminder(4)) Then
                                blnTriggerWhat = True
                            End If
                        Else
                            frmMain.lstReminders.ListItems(X).SubItems(2) = Date
                        End If
                    End If
                    If (CLng(arrReminder(2)) And enmReminderWhen.REMIND_EVERY_YEAR_AT_SPECIFIED_DATE_AND_TIME) = enmReminderWhen.REMIND_EVERY_YEAR_AT_SPECIFIED_DATE_AND_TIME Then
                        If (Day(Date) = Day(CDate(arrReminder(3)))) And (Month(Date) = Month(CDate(arrReminder(3)))) Then
                            If Time >= CDate(arrReminder(4)) Then
                                blnTriggerWhat = True
                            End If
                        Else
                            frmMain.lstReminders.ListItems(X).SubItems(2) = Date
                        End If
                    End If
                    If (CLng(arrReminder(2)) And enmReminderWhen.REMIND_EVERYDAY_AT_SPECIFIED_TIME) = enmReminderWhen.REMIND_EVERYDAY_AT_SPECIFIED_TIME Then
                        If Time >= CDate(arrReminder(4)) Then
                            blnTriggerWhat = True
                        End If
                    End If
                    If (CLng(arrReminder(2)) And enmReminderWhen.REMIND_EVERY_MONTH_AT_SPECIFIED_DATE_AND_TIME) = enmReminderWhen.REMIND_EVERY_MONTH_AT_SPECIFIED_DATE_AND_TIME Then
                        If (Day(Date) = Day(CDate(arrReminder(3)))) Then
                            If Time > CDate(arrReminder(4)) Then
                                blnTriggerWhat = True
                            End If
                        Else
                            frmMain.lstReminders.ListItems(X).SubItems(2) = Date
                        End If
                    End If
                    
                    If blnTriggerWhat Then
                        If (CLng(arrReminder(5)) And enmReminderWhat.PLAY_SOUND) = enmReminderWhat.PLAY_SOUND Then
                            sndPlaySound arrReminder(7), &H1
                        End If
                        If (CLng(arrReminder(5)) And enmReminderWhat.SHOW_MESSAGE) = enmReminderWhat.SHOW_MESSAGE Then
                            showMessage arrReminder(6), arrReminder(1)
                        End If
                        
                        frmMain.lstReminders.ListItems(X).SubItems(2) = Date
                    End If
                
                End If
            End If
        End If
        
    Next
End Sub

' stores the reminders from the listview to the registry & returns the
' current string of reminders that should be used at position 0,
' and the lastrun dates in position 1
Public Sub remindersSave()
    Dim objReg As clsRegistry
    Dim strReminders As String, X As Long
    
    strReminders = ""
    For X = 1 To frmMain.lstReminders.ListItems.Count
        strReminders = strReminders & frmMain.lstReminders.ListItems(X).Tag
        strReminders = strReminders & frmMain.lstReminders.ListItems(X).SubItems(2) & "*"
    Next
    
    Set objReg = New clsRegistry
    With objReg
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = m_strRegRoot
        
        .ValueKey = "m_strReminders"
        .Value = strReminders
    End With
    Set objReg = Nothing
End Sub

' retreives the reminders from the registry, and stores them into the
' listview, also returns the string read from the registry
Public Sub remindersGet()
    Dim objReg As clsRegistry
    Dim strReminders As String
    Dim arrReminders() As String, arrReminder() As String, X As Long
    
    Set objReg = New clsRegistry
    With objReg
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = m_strRegRoot
        
        .ValueKey = "m_strReminders"
        strReminders = .Value
    End With
    Set objReg = Nothing
    
    arrReminders = Split(strReminders, "*")
    
    frmMain.lstReminders.ListItems.Clear
    For X = LBound(arrReminders) To UBound(arrReminders)
        arrReminder = Split(arrReminders(X), "|")
        If UBound(arrReminder) = 8 Then
            If arrReminder(0) = "1.0" Then
                With frmMain.lstReminders.ListItems.Add(, , arrReminder(3) & " " & arrReminder(4))
                    .SubItems(1) = arrReminder(1)
                    .Tag = Left(arrReminders(X), InStrRev(arrReminders(X), "|"))
                    .SubItems(2) = arrReminder(8)
                End With
            End If
        End If
    Next
End Sub
