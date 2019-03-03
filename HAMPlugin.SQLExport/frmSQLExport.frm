VERSION 5.00
Begin VB.Form frmSQLExport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export to SQL Server"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSQLExport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   287
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   428
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   3735
      TabIndex        =   8
      Top             =   3690
      Width           =   1125
   End
   Begin VB.CommandButton btnExport 
      Caption         =   "&Export"
      Height          =   345
      Left            =   4965
      TabIndex        =   7
      Top             =   3690
      Width           =   1125
   End
   Begin VB.TextBox txtUsername 
      Height          =   300
      Left            =   375
      MaxLength       =   50
      TabIndex        =   6
      Top             =   2310
      Width           =   5340
   End
   Begin VB.CommandButton cmdConnectionString 
      Caption         =   ".."
      Height          =   300
      Left            =   5775
      TabIndex        =   4
      Top             =   1410
      Width           =   315
   End
   Begin VB.TextBox txtConnectionString 
      Height          =   300
      Left            =   375
      TabIndex        =   3
      Top             =   1410
      Width           =   5340
   End
   Begin VB.Label lblUsername 
      AutoSize        =   -1  'True
      Caption         =   "Username (to uniquely identify within your organization):"
      Height          =   195
      Left            =   375
      TabIndex        =   5
      Top             =   2085
      Width           =   4095
   End
   Begin VB.Label lblConnectionString 
      AutoSize        =   -1  'True
      Caption         =   "Connection String:"
      Height          =   195
      Left            =   375
      TabIndex        =   2
      Top             =   1185
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   437
      Y1              =   59
      Y2              =   59
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSQLExport.frx":000C
      Height          =   495
      Left            =   375
      TabIndex        =   1
      Top             =   345
      Width           =   5760
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Export to SQL Server Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   210
      TabIndex        =   0
      Top             =   135
      Width           =   2505
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   930
      Left            =   0
      Top             =   0
      Width           =   6570
   End
End
Attribute VB_Name = "frmSQLExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_blnExportSelected As Boolean
Public m_blnClosed As Boolean

Public Sub ShowMe()
    txtConnectionString.Text = GetSetting("HAMPlugin.SQLExport", "Settings", "ConnectionString", "")
    txtUsername.Text = GetSetting("HAMPlugin.SQLExport", "Settings", "Username", "")
    m_blnExportSelected = False
    m_blnClosed = False
    Show
    AppActivate Caption
End Sub

Private Sub btnCancel_Click()
    m_blnClosed = True
    Hide
End Sub

Private Sub btnExport_Click()
    Const adStateOpen = 1
    Dim strErrors As String
    Dim objADO As Object
    
    On Error Resume Next
    strErrors = ""
    If Len(Trim(txtUsername.Text)) = 0 Then
        strErrors = strErrors & "- Username field can not be empty." & vbCrLf
    End If
    If Len(Trim(txtConnectionString.Text)) = 0 Then
        strErrors = strErrors & "- Connection string field can not be empty." & vbCrLf
    Else
        Set objADO = CreateObject("ADODB.Connection")
        If objADO Is Nothing Then
            strErrors = strErrors & "- Error creating ADODB.Connection object, please install MDAC." & vbCrLf
        Else
            objADO.Open txtConnectionString.Text
            If objADO.State <> adStateOpen Then
                strErrors = strErrors & "- Error connecting to datasource, make sure that the ""Alow saving password"" checkbox is checked." & vbCrLf
            Else
                objADO.Close
            End If
            Set objADO = Nothing
        End If
    End If
    
    If Len(strErrors) > 0 Then
        MsgBox "The export could not start since the following error(s) was raised:" & vbCrLf & vbCrLf & strErrors, vbCritical Or vbOKOnly
        Exit Sub
    End If
    
    ' save settings for use next round
    SaveSetting "HAMPlugin.SQLExport", "Settings", "ConnectionString", txtConnectionString.Text
    SaveSetting "HAMPlugin.SQLExport", "Settings", "Username", txtUsername.Text

    m_blnExportSelected = True
    m_blnClosed = True
    Hide
End Sub

Private Sub cmdConnectionString_Click()
    Dim objDatalinks As Object
    Dim strConnectionString As String
    
    On Error GoTo lblError
    Set objDatalinks = CreateObject("Datalinks")
    
    On Error Resume Next
    strConnectionString = objDatalinks.PromptNew
    Set objDatalinks = Nothing
    If strConnectionString <> "" Then
        txtConnectionString.Text = strConnectionString
    End If
    
    Exit Sub
lblError:
    MsgBox "There was an error creating the Datalinks connection string object." & vbCrLf & vbCrLf & "You'll have to type the connection string in manually.", vbOKOnly Or vbCritical
    Exit Sub
End Sub
