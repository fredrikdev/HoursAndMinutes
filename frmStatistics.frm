VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmStatistics 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Statistics"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStatistics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   334
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   461
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgPrinter 
      Left            =   150
      Top             =   4395
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton btnPrint 
      Caption         =   "&Print..."
      Height          =   345
      Left            =   4395
      TabIndex        =   1
      Top             =   4500
      Width           =   1125
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   5640
      TabIndex        =   2
      Top             =   4500
      Width           =   1125
   End
   Begin VB.TextBox txtReport 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4185
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   30
      Width           =   6915
   End
   Begin VB.Frame frmBottom 
      Height          =   120
      Left            =   -135
      TabIndex        =   4
      Top             =   4125
      Width           =   7290
   End
   Begin VB.Frame frmTop 
      Height          =   120
      Left            =   -60
      TabIndex        =   3
      Top             =   -90
      Width           =   7155
   End
End
Attribute VB_Name = "frmStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub CreateReport(lngMode As Long, dteStart As Date, dteEnd As Date)
    Dim objReg As clsRegistry
    Dim dteCurrent As Date, lngCurrent As Long
    Dim x As Long, y As Long, arrComments() As String
    Dim lngCTime As Long, lngTTime As Long
    Dim lngSTTime As Long, lngModesCalc As Long
    
    ' initialize
    lngSTTime = 0
    lngModesCalc = 0
    txtReport.Text = ""
    
    For x = 1 To m_colModes.Count
        lngCTime = 0
        lngTTime = 0
        dteCurrent = dteStart
        lngCurrent = CLng(dteStart)
        
        If (m_colModes.Item(x).m_lngID = lngMode) Or (lngMode = -1) Then
            lngModesCalc = lngModesCalc + 1
            
            ' write report header
            txtReport.Text = txtReport.Text & "Time spent on '" & m_colModes.Item(x).m_strName & "' from " & FormatDateTime(dteStart, vbShortDate) & " to " & FormatDateTime(dteEnd, vbShortDate) & "." & vbCrLf & vbCrLf
            
            ' write report days
            Set objReg = New clsRegistry
            objReg.ClassKey = HKEY_CURRENT_USER
            While (dteCurrent <= dteEnd)
                objReg.SectionKey = m_strRegRoot & "\Modes\" & m_colModes.Item(x).m_lngID
                objReg.ValueKey = "Date " & lngCurrent
                lngCTime = objReg.Value
                objReg.ValueKey = "Comment " & lngCurrent
                arrComments = Split(CStr(objReg.Value), "*")
                lngTTime = lngTTime + lngCTime
                txtReport.Text = txtReport.Text & FormatDateTime(dteCurrent, vbShortDate) & ":  " & FormatMSeconds(lngCTime) & vbCrLf
                For y = LBound(arrComments) To UBound(arrComments)
                    If Trim(arrComments(y)) <> "" Then
                        txtReport = txtReport.Text & "             " & arrComments(y) & vbCrLf
                        If y = UBound(arrComments) Then txtReport = txtReport.Text & vbCrLf
                    End If
                Next
                dteCurrent = DateAdd("d", 1, dteCurrent)
                lngCurrent = CLng(dteCurrent)
            Wend
            Set objReg = Nothing
    
            ' write report bottom
            txtReport.Text = txtReport.Text & vbCrLf & "Total Time:  " & FormatMSeconds(lngTTime) & vbCrLf & vbCrLf & vbCrLf
            
            lngSTTime = lngSTTime + lngTTime
        End If
    Next
    If lngModesCalc > 1 Then
        txtReport.Text = txtReport.Text & "Total Time:  " & FormatMSeconds(lngSTTime) & " (all modes)"
    End If
End Sub

Private Sub btnOk_Click()
    Hide
End Sub

Private Sub btnPrint_Click()
    On Error GoTo lblError
    dlgPrinter.PrinterDefault = True
    dlgPrinter.ShowPrinter
    
    Printer.Font = txtReport.Font
    Printer.Copies = dlgPrinter.Copies
    Printer.Orientation = dlgPrinter.Orientation
    Printer.CurrentX = Printer.TwipsPerPixelX * 100
    Printer.CurrentY = Printer.TwipsPerPixelY * 100
    
    Printer.Print Replace(vbCrLf & vbCrLf & App.Title & " " & Now & vbCrLf & vbCrLf & txtReport.Text, vbCrLf, vbCrLf & vbTab)
    Printer.EndDoc
lblError:
End Sub

