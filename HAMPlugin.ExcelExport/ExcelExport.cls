VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExcelExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =============================================================================
'
' Hours and Minutes "Export to Excel" Plugin
'
' This Plugin is provided with source to aid you in how to create a custom
' statistics plugin for Hours and Minutes.
'
' Some general Rules:
'
' -  The compiled DLL must be named according as follow:
'    <projectname>.<classname>.phm
'    The .phm extension is associated with the Plugin Installer (PI), providing
'    doubleclick-to-install capability. The naming <projectname>.<classname>
'    is for the CreateObject part.
' -  The compiled DLL must include the following "get" properties:
'    PluginType : "Hours and Minutes Statistics Plugin 1.0"
'    PluginCaption : "Export to Excel"
'    And the following Function:
'    CreateReport(arrMode() AS HAMPluginLib.clsMode
'
' The CreateReport sub and the arrMode() array:
'
' - The CreateReport sub is called when a user of Hours and Minutes, clicks
'   on the menu associated with your custom plugin.
' - As a parameter to the function is the data that the user has selected, it's
'   provided named arrMode.
' - arrMode is an array of the mode(s) that the user has selected.
'   ie. arrMode(0).modeName is the name of the first mode, and so on, use
'   "for each" or a "for x = lbound(arrMode) to ubound(arrMode)" to loop the
'   modes through.
' - in each item in arrMode (ie arrMode(5)) you may access each day that
'   has been selected. arrMode(x).DDate is the date, arrMode(x).Milliseconds is
'   the number of milliseconds for the date, arrMode(x).Comments() is a string
'   array for all the comments made the date (again: use a "for each" or a
'   "for x = lbound(arrMode(x).Comments) to ubound(arrMode(x).Comments)" to
'   loop the comments).
' =============================================================================

Option Explicit

' This property must always be present in order for Hours and Minutes
' to identify the Plugin correctly
Public Property Get PluginType() As String
    PluginType = "Hours and Minutes Statistics Plugin 1.0"
End Property

' This property must always be present in order for Hours and Minutes PI
' to determine the caption for the menuitem representing the Plugin
Public Property Get PluginCaption() As String
    PluginCaption = "Export to Excel"
End Property

' This sub is called by Hours and Minutes with an array containing
' all data to be in the report
Public Sub CreateReport(ByRef arrMode() As HAMPluginLib.clsMode)
    Dim objExcel As Excel.Application, objWS As Excel.Worksheet
    Dim intCol As Integer, intRow As Integer
    Dim objMode As Variant, objDay As Variant
    Dim blnListDates As Boolean, lngCTime As Long, lngTTime As Long
    Dim strTemp As Variant
    
    ' initialize
    blnListDates = True
    
    ' create excel
    Set objExcel = New Excel.Application
    objExcel.Visible = True
    Set objWS = objExcel.Workbooks.Add.ActiveSheet
    
    ' format
    objWS.Cells.Font.Name = "Tahoma"
    objWS.Cells.Font.Size = 9
    objWS.Cells.ColumnWidth = 10
    
    ' write header
    objWS.Cells(1, 1) = "Hours and Minutes - Excel Export"
    objWS.Cells(1, 1).Font.Size = 13
    objWS.Cells(1, 1).Font.Bold = True
    objExcel.StatusBar = "Please wait while the report is beeing created..."
    
    ' create report
    intCol = 2
    lngTTime = 0
    For Each objMode In arrMode
        ' print out the mode name
        objWS.Cells(3, intCol) = objMode.modeName
        objWS.Cells(3, intCol).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
        objWS.Cells(3, intCol).ColumnWidth = 15
        intRow = 4
        
        ' print out time of each date
        lngCTime = 0
        For Each objDay In objMode.modeDays
            If blnListDates Then
                objWS.Cells(intRow, 1) = objDay.DDate
                objWS.Cells(intRow, 1).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                objWS.Cells(intRow, 1).VerticalAlignment = Excel.XlVAlign.xlVAlignTop
            End If
            objWS.Cells(intRow, intCol) = " " & FormatMSeconds(objDay.Milliseconds)
            objWS.Cells(intRow, intCol).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
            objWS.Cells(intRow, intCol).VerticalAlignment = Excel.XlVAlign.xlVAlignTop
            For Each strTemp In objDay.Comments
                strTemp = CStr(strTemp)
                If strTemp <> "" Then
                    objWS.Cells(intRow, intCol) = objWS.Cells(intRow, intCol) & vbLf & strTemp
                End If
            Next
            lngCTime = lngCTime + objDay.Milliseconds
            intRow = intRow + 1
        Next
        
        ' calculate column totals
        If blnListDates Then
            objWS.Cells(intRow + 2, 1) = "Row Total:"
        End If
        objWS.Cells(intRow + 2, intCol) = " " & FormatMSeconds(lngCTime)
        objWS.Cells(intRow + 2, intCol).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
        lngTTime = lngTTime + lngCTime
        
        intCol = intCol + 1
        blnListDates = False
    Next
    
    ' calculate totals
    objWS.Cells(intRow + 4, 1) = "Total Time:"
    objWS.Cells(intRow + 4, 2) = " " & FormatMSeconds(lngTTime)
    objWS.Cells(intRow + 4, 2).HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
    
    objExcel.StatusBar = ""
    Set objWS = Nothing
    Set objExcel = Nothing
End Sub

' format milliseconds as "HH:MM:SS"
Private Function FormatMSeconds(ByVal lngMSeconds As Long)
    Dim lngH As Long, lngM As Long
    
    lngMSeconds = lngMSeconds \ 1000
    
    lngH = lngMSeconds \ 60 \ 60
    lngMSeconds = lngMSeconds - (lngH * 60 * 60)
    
    lngM = lngMSeconds \ 60
    lngMSeconds = lngMSeconds - (lngM * 60)
    
    FormatMSeconds = IIf(lngH < 10, "0" & lngH, lngH) & ":" & IIf(lngM < 10, "0" & lngM, lngM) & ":" & IIf(lngMSeconds < 10, "0" & lngMSeconds, lngMSeconds)
End Function
