Attribute VB_Name = "modRegCodeAlgorithm"
Option Explicit

Public m_blnRegistered As Boolean
Public m_lngStartCount As Long

Function regCodeEvaluate(strCode As String) As String
    Const strMasterMask = "9EAZD1K70X24YR52V3C32HN4JI7G8051FBMLO5P61QS7T02U556WQW49FAS423MC"
    Const strUserMask = "0f129874lhaf25dqwerAsdASFPHWE12463049dsajxzpoq5213rvsb23gweqwet2"
    
    Dim arrTemp() As String
    Dim strUser As String, x As Long, y As Long
    Dim strHash As String
    
    ' initialize error handler
    On Error GoTo lblError
    
    ' first create the regcode using the generator algorithm
    arrTemp = Split(strCode, vbCrLf)
    If (LBound(arrTemp) <> 0) Or (UBound(arrTemp) <> 8) Then Err.Raise 1000
    If Trim(arrTemp(0)) <> "-- BEGIN LICENSE --" Then Err.Raise 1000
    If Trim(arrTemp(8)) <> "-- END LICENSE   --" Then Err.Raise 1000
    
    ' check that we have the same major version
    If App.Major > CInt(Mid$(arrTemp(1), 19, 1)) Then Err.Raise 1001

    ' create strUser string
    strUser = Trim(arrTemp(2)) & " " & Trim(arrTemp(3)) & " " & Trim(arrTemp(4)) & " " & Trim(Mid$(arrTemp(1), 19))
    x = 0
    y = charAt(strUser, 0)
    Do While (Len(strUser) < 80)
        strUser = strUser & Chr(charAt(strUserMask, y Mod Len(strUserMask)))
        y = charAt(strUser, y Mod Len(strUser))
        x = x + 3
        y = y + x
    Loop
    strUser = Left(strUser, 72)
    
    ' create hash
    If doHash(strUser, strMasterMask) <> Trim(arrTemp(5)) & Trim(arrTemp(6)) & Trim(arrTemp(7)) Then Err.Raise 1000
    
    regCodeEvaluate = "OK"
    m_blnRegistered = True
    Exit Function
lblError:
    Select Case Err.Number
        Case 1000: regCodeEvaluate = "Invalid License Format." & vbCrLf & vbCrLf & "Remember to include both the -- BEGIN LICENSE -- and the -- END LICENSE   -- row!"
        Case 1001: regCodeEvaluate = "This licese is only valid for an older version of Hours and Minutes"
    End Select
    m_blnRegistered = False
    Exit Function
End Function

Function charAt(strString As String, lngPos As Long) As Long
    charAt = CLng(Asc(Mid$(strString, lngPos + 1, 1)))
End Function

Function doHash(strString As String, strMask As String) As String
    Dim strResult01 As String, strResult02 As String, strResult03 As String
    Dim strResult As String
    Dim x As Long, y As Long, z As Long
    Dim bytChar As Byte, bytHashChar As Byte
    
    ' pass 1 (hash jump encoding)
    For y = 0 To Len(strString) - 1
        x = charAt(strString, y)
        x = x Mod Len(strString)
        bytChar = charAt(strString, x)
        strResult01 = strResult01 & Chr(charAt(strMask, Abs(bytChar - Len(strString)) Mod Len(strMask)))
    Next
    
    ' pass 2 (lame encoding)
    For y = 0 To Len(strString) - 1
        bytChar = charAt(strString, y)
        strResult02 = strResult02 & Chr(charAt(strMask, bytChar Mod Len(strMask)))
    Next
    
    ' pass 3 (combine strings)
    x = 0
    y = 0
    Do While (x < Len(strResult01)) Or (y < Len(strResult02))
        If (x < Len(strResult01)) Then strResult03 = strResult03 & Chr(charAt(strResult01, x))
        If (y < Len(strResult02)) Then strResult03 = strResult03 & Chr(charAt(strResult02, y))
        x = x + 1
        y = y + 1
    Loop
    
    ' pass 4 (even split - not used in vb)
    doHash = strResult03
End Function
