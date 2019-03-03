Attribute VB_Name = "modFunctions"
Option Explicit

' declarations
Public Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Any, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE

' constants:
Public Const m_strRegRoot = "SOFTWARE\Port Jackson Computing\Hours and Minutes"

' format milliseconds as "HH:MM:SS"
Public Function FormatMSeconds(ByVal lngMSeconds As Long)
    Dim lngH As Long, lngM As Long
    
    lngMSeconds = lngMSeconds \ 1000
    
    lngH = lngMSeconds \ 60 \ 60
    lngMSeconds = lngMSeconds - (lngH * 60 * 60)
    
    lngM = lngMSeconds \ 60
    lngMSeconds = lngMSeconds - (lngM * 60)
    
    FormatMSeconds = IIf(lngH < 10, "0" & lngH, lngH) & ":" & IIf(lngM < 10, "0" & lngM, lngM) & ":" & IIf(lngMSeconds < 10, "0" & lngMSeconds, lngMSeconds)
End Function

' get milliseconds from a time
Public Function GetMSeconds(ByVal dteDate As Date) As Long
    GetMSeconds = DatePart("s", dteDate) * 1000
    GetMSeconds = GetMSeconds + DatePart("n", dteDate) * 60 * 1000
    GetMSeconds = GetMSeconds + DatePart("h", dteDate) * 60 * 60 * 1000
End Function

' checks if a value is a integer value
Public Function isInteger(ByVal varValue As Variant) As Boolean
    Dim X As Integer, strTemp As String
    
    isInteger = False
    strTemp = CStr(varValue)
    If Len(strTemp) = 0 Then Exit Function
    
    For X = 1 To Len(strTemp)
        If Mid$(strTemp, X, 1) < "0" Or Mid$(strTemp, X, 1) > "9" Then Exit Function
    Next
    
    isInteger = True
End Function

' gets a input text from the user
Public Function getInput(objOwnerForm As Form, strPrompt As String, strTitle As String, strDefault As String) As String
    getInput = ""
    
    Load frmInput
    With frmInput
        .lblPrompt.Caption = strPrompt
        .Caption = strTitle
        .txtValue.Text = strDefault
        .txtValue.SelStart = 0
        .txtValue.SelLength = Len(.txtValue.Text)
        .Show 1, objOwnerForm
        If .m_blnCancelPressed = False Then getInput = Trim(.txtValue.Text)
    End With
    Unload frmInput
End Function

' shows a message to the user
Public Sub showMessage(strMessage As String, strTitle As String)
    Dim X As Long
    Dim blnShowMessagesModally As Boolean
    
    blnShowMessagesModally = False
    If Forms.Count > 1 Then blnShowMessagesModally = True

    On Error Resume Next
    With frmMessage
        .AddMessage strMessage, strTitle
        If blnShowMessagesModally = False Then
            .Show 0
        Else
            .Show 1
        End If
        .SetFocus
    End With
    On Error GoTo 0
End Sub

' checks if a string (filename) contain invalid characters
Public Function containInvalidChars(strText As String) As Boolean
    Dim strTemp As String, X As Long

    containInvalidChars = True

    For X = 1 To Len(strText)
        strTemp = Mid$(strText, X, 1)
        If strTemp = "\" Or strTemp = "/" Or strTemp = ":" Or strTemp = "*" Or strTemp = "?" Or strTemp = """" Or strTemp = "<" Or strTemp = ">" Or strTemp = "|" Then Exit Function
    Next
    
    containInvalidChars = False
End Function

' returns the name of the weekday for a date
Public Function DayName(dteDate As Date) As String
    Select Case Weekday(dteDate)
        Case 1: DayName = "Sunday"
        Case 2: DayName = "Monday"
        Case 3: DayName = "Tuesday"
        Case 4: DayName = "Wednesday"
        Case 5: DayName = "Thursday"
        Case 6: DayName = "Friday"
        Case 7: DayName = "Saturday"
    End Select
End Function

' checks if a index exist within a collection
Public Function isInCollection(ByVal strIndex As String, ByRef colCollection As Collection) As Boolean
    On Error GoTo lblErrorHandler
    isInCollection = True
    With colCollection.Item(strIndex)
    End With
    On Error GoTo 0
    Exit Function
lblErrorHandler:
    On Error GoTo 0
    isInCollection = False
End Function

' returns the idle time in seconds (if called once every second)
Public Sub getIdleTime(lngTCount As Long, dteDate As Date, ByRef lngIdleTimeTotal As Long, ByRef lngIdleTimeToday As Long)
    Static objOldPoint As POINTAPI
    Static objLastRun As Variant
    Static objLastDate As Variant, objLastRunDate As Variant
    Dim objPoint As POINTAPI, X As Long, blnTemp As Boolean
    lngIdleTimeTotal = 0
    lngIdleTimeToday = 0
    
    ' check if cursor has been moved since last run
    GetCursorPos objPoint
    
    blnTemp = False
    If objPoint.X <> objOldPoint.X Or objPoint.Y <> objOldPoint.Y Then
        blnTemp = True
    End If
    For X = 1 To 255
        If GetAsyncKeyState(X) <> 0 Then
            blnTemp = True
        End If
    Next
    
    If blnTemp = False Then
        ' idle
        If IsEmpty(objLastRun) Then objLastRun = lngTCount
        If IsEmpty(objLastDate) Then
            objLastDate = dteDate
            objLastRunDate = lngTCount
        End If
        If objLastDate <> dteDate Then
            objLastDate = dteDate
            objLastRunDate = lngTCount
        End If
        lngIdleTimeToday = lngTCount - objLastRunDate
        lngIdleTimeTotal = lngTCount - objLastRun
    Else
        ' active
        objOldPoint = objPoint
        objLastDate = Empty
        objLastRun = Empty
        objLastRunDate = Empty
    End If
End Sub

' repeats a string # number of times
Public Function Repeat(strString As String, lngCount As Long) As String
    Dim X As Long
    Repeat = ""
    For X = 1 To lngCount
        Repeat = Repeat & strString
    Next
End Function

' trims and includes crlf characters
Public Function TrimCRLF(strString As String) As String
    While Right$(strString, 1) = vbCr Or Right$(strString, 1) = vbLf
        strString = Trim$(Left$(strString, Len(strString) - 1))
    Wend
    While Left$(strString, 1) = vbCr Or Left$(strString, 1) = vbLf
        strString = Trim$(Right$(strString, Len(strString) - 1))
    Wend
    TrimCRLF = Trim$(strString)
End Function

