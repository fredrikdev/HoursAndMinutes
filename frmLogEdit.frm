VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLogEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Log"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dteDate 
      Height          =   315
      Left            =   4890
      TabIndex        =   3
      Top             =   1035
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      _Version        =   393216
      Format          =   19791873
      UpDown          =   -1  'True
      CurrentDate     =   37112
   End
   Begin VB.ComboBox lstModes 
      Height          =   315
      Left            =   165
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1035
      Width           =   4485
   End
   Begin MSComCtl2.DTPicker dteTime 
      Height          =   330
      Left            =   165
      TabIndex        =   5
      Top             =   1995
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      _Version        =   393216
      Format          =   19791874
      UpDown          =   -1  'True
      CurrentDate     =   37112
   End
   Begin HoursAndMinutes.ctlCommentEdit ctlComments 
      Height          =   1470
      Left            =   1785
      TabIndex        =   7
      Top             =   1995
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   2593
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   5130
      TabIndex        =   9
      Top             =   3900
      Width           =   1125
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   345
      Left            =   3915
      TabIndex        =   8
      Top             =   3900
      Width           =   1125
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Select a &date:"
      Height          =   195
      Left            =   4890
      TabIndex        =   2
      Top             =   825
      Width           =   1005
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "&Select a task:"
      Height          =   195
      Left            =   165
      TabIndex        =   0
      Top             =   825
      Width           =   975
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "&Comments:"
      Height          =   195
      Left            =   1785
      TabIndex        =   6
      Top             =   1785
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Time:"
      Height          =   195
      Left            =   165
      TabIndex        =   4
      Top             =   1785
      Width           =   390
   End
   Begin VB.Label Label1 
      Caption         =   "Select the task and date that you're interested in from the list && box below. You'll be asked to save any changes."
      Height          =   495
      Left            =   165
      TabIndex        =   10
      Top             =   150
      Width           =   6090
   End
End
Attribute VB_Name = "frmLogEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_lngFocusModeID As Long
Private m_blnChanged As Boolean

Private m_lngCurrentMode As Long
Private m_dteCurrentDate As Date

Public m_blnCancelPressed As Boolean

Public Sub listModes()
    Dim X As Long
    lstModes.Clear
    
    For X = 1 To m_colModes.Count
        With m_colModes.Item(X)
            lstModes.AddItem .m_strName
            lstModes.ItemData(lstModes.ListCount - 1) = .m_lngID
            If m_lngFocusModeID = .m_lngID Then
                lstModes.ListIndex = lstModes.ListCount - 1
            End If
        End With
    Next
    
    GetData
End Sub

Private Sub saveCurrent()
    Dim objReg As clsRegistry
    Dim arrComments() As String, strComments As String
    Dim strTemp As String

    If m_blnChanged Then
        strTemp = ""
        If (m_lngCurrentMode = frmMain.m_lngActiveModeID) And (m_dteCurrentDate = frmMain.m_dteDate) Then
            strTemp = vbCrLf & vbCrLf & "Note: Since this is the current task & date used, time will also be updated to the counter."
        End If
    
        If MsgBox("Do you want to save changes to the current date?" & strTemp, vbYesNo Or vbQuestion) = vbYes Then
            arrComments = ctlComments.GetText
            strComments = Join(arrComments, "*")
            strComments = Left(strComments, InStrRev(strComments, "*"))
            
            ' check if current mode
            If (m_lngCurrentMode = frmMain.m_lngActiveModeID) And (m_dteCurrentDate = frmMain.m_dteDate) Then
                frmMain.m_lngStartTime = GetTickCount - GetMSeconds(dteTime.Value)
            End If
            
            ' save to registry
            Set objReg = New clsRegistry
            With objReg
                .ClassKey = HKEY_CURRENT_USER
                .SectionKey = m_strRegRoot & "\Modes\" & m_lngCurrentMode
                .ValueKey = "Date " & CLng(m_dteCurrentDate)
                .Value = GetMSeconds(dteTime.Value)
                .ValueKey = "Comment " & CLng(m_dteCurrentDate)
                .Value = strComments
            End With
            Set objReg = Nothing
        End If
        m_blnChanged = False
    End If
End Sub

Private Sub GetData()
    Dim objReg As clsRegistry
    Dim strComments As String, arrComments() As String
    
    ' check for changes to the current post
    saveCurrent
    
    Set objReg = New clsRegistry
    With objReg
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = m_strRegRoot & "\Modes\" & lstModes.ItemData(lstModes.ListIndex)
    
        .ValueKey = "Date " & CLng(dteDate.Value)
        dteTime.Value = FormatMSeconds(CLng(.Value))
        
        .ValueKey = "Comment " & CLng(dteDate.Value)
        strComments = CStr(.Value)
        
        If strComments <> "" Then
            strComments = strComments & "<new comment>"
            arrComments = Split(strComments, "*")
            ctlComments.SetText arrComments
        Else
            arrComments = Split("<new comment>", "*")
            ctlComments.SetText arrComments
        End If
    End With
    Set objReg = Nothing
    
    m_lngCurrentMode = lstModes.ItemData(lstModes.ListIndex)
    m_dteCurrentDate = CDate(dteDate.Value)
End Sub

Private Sub btnCancel_Click()
    Hide
End Sub

Private Sub btnOk_Click()
    m_blnCancelPressed = False
    saveCurrent
    Hide
End Sub

Private Sub ctlComments_DblClick(strText As String, blnUpdate As Boolean, blnAddItem As Boolean)
    If strText = "<new comment>" Then
        Load frmModeComment
        With frmModeComment
            .txtValue.Text = ""
            .Show 1, Me
            If .m_blnCancelPressed = False Then
                strText = .txtValue.Text
                blnAddItem = True
                m_blnChanged = True
            End If
        End With
        Unload frmModeComment
    Else
        Load frmModeComment
        With frmModeComment
            .txtValue.Text = strText
            .txtValue.SelStart = 0
            .txtValue.SelLength = Len(strText)
            .Show 1, Me
            If .m_blnCancelPressed = False Then
                strText = .txtValue.Text
                blnUpdate = True
                m_blnChanged = True
            End If
        End With
        Unload frmModeComment
    End If
End Sub

Private Sub ctlComments_Delete(strText As String, blnDelete As Boolean)
    If strText = "<new comment>" Then Exit Sub
    If MsgBox("Are you sure that you want to delete the selected comment?", vbYesNo Or vbQuestion) = vbYes Then
        blnDelete = True
        m_blnChanged = True
    End If
End Sub

Private Sub dteDate_Change()
    GetData
End Sub

Private Sub dteTime_Change()
    m_blnChanged = True
End Sub

Private Sub Form_Load()
    m_blnCancelPressed = True
    dteDate.Value = Date
    dteTime.Value = CDate(Date)
    m_blnChanged = False
End Sub

Private Sub lstModes_Click()
    GetData
End Sub

