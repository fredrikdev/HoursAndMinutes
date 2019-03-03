VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SQLExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' =============================================================================
'
' Hours and Minutes "Export to SQL Server" Plugin
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
'    PluginCaption : "Export to SQL"
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
    PluginCaption = "Export to SQL Server"
End Property

' This sub is called by Hours and Minutes with an array containing
' all data to be in the report
Public Sub CreateReport(ByRef arrMode() As HAMPluginLib.clsMode)
    Dim objADO As Object, strUsername As String, strConnectionString As String
    Dim objRS As Object
    Dim objTask As Variant, objDay As Variant, strComment As Variant
    Dim strTemp As String
    Dim strSQL As String
    
    ' get settings
    Load frmSQLExport
    frmSQLExport.ShowMe
    While Not frmSQLExport.m_blnClosed
        DoEvents
    Wend
    If Not frmSQLExport.m_blnExportSelected Then
        Unload frmSQLExport
        Exit Sub
    End If
    strUsername = Replace(frmSQLExport.txtUsername.Text, "'", "''")
    strConnectionString = frmSQLExport.txtConnectionString.Text
    Unload frmSQLExport
    
    ' connect to datasource
    Set objADO = CreateObject("ADODB.Connection")
    objADO.Open strConnectionString
    
    ' do a create table (this will create if table not exist)
    objADO.Execute "" & _
    "IF NOT EXISTS (SELECT name FROM sysobjects WHERE name = 'tblHAMTimedata') BEGIN " & vbCrLf & _
        "CREATE TABLE [dbo].[tblHAMTimedata] (" & _
        "[ID] [int] IDENTITY (1, 1) NOT NULL, " & _
        "[Username] [varchar] (50) NOT NULL, " & _
        "[Date] [datetime] NOT NULL, " & _
        "[Task] [varchar] (50) NOT NULL, " & _
        "[Comments] [varchar] (255) NOT NULL, " & _
        "[Seconds] [numeric](18, 0) NULL " & _
        ") ON [PRIMARY] " & vbCrLf & _
        "ALTER TABLE [dbo].[tblHAMTimedata] WITH NOCHECK ADD " & _
        "CONSTRAINT [PK_tblHAMTimedata] PRIMARY KEY CLUSTERED ([ID]) ON [PRIMARY] " & vbCrLf & _
        "CREATE  INDEX [IX_tblHAMTimedata] ON [dbo].[tblHAMTimedata]([Username]) ON [PRIMARY]" & vbCrLf & _
    "END"
    
    ' insert/update selected data
    For Each objTask In arrMode
        For Each objDay In objTask.modeDays
            ' gather comments
            strTemp = ""
            For Each strComment In objDay.Comments
                strTemp = strTemp & strComment & vbCrLf
            Next
            strTemp = Replace(Left(strTemp, 255), "'", "''")
            
            strSQL = "SELECT ID FROM tblHAMTimedata WHERE Username='" & strUsername & "' AND [Date]='" & objDay.DDate & "' AND Task='" & Replace(objTask.modeName, "'", "''") & "'"
            If objADO.Execute(strSQL).EOF Then
                ' do an insert
                strSQL = "INSERT tblHAMTimedata (Username, [Date], Task, Comments, Seconds) VALUES ('" & strUsername & "','" & objDay.DDate & "','" & Replace(objTask.modeName, "'", "''") & "',"
                strSQL = strSQL & "'" & strTemp & "'," & CLng(objDay.Milliseconds \ 1000) & ")"
                objADO.Execute strSQL
            Else
                ' do an update
                strSQL = "UPDATE tblHAMTimedata SET Comments='" & strTemp & "', Seconds=" & CLng(objDay.Milliseconds \ 1000) & " "
                strSQL = strSQL & "WHERE Username='" & strUsername & "' AND [Date]='" & objDay.DDate & "' AND Task='" & Replace(objTask.modeName, "'", "''") & "'"
                objADO.Execute strSQL
            End If
        Next
    Next
    
    objADO.Close
    Set objADO = Nothing
    MsgBox "Data was Successfully Exported to SQL Server", vbInformation Or vbOKOnly
End Sub
