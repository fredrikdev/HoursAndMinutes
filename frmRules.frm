VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmRules 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rules"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRules.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   385
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   404
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnTest 
      Caption         =   "&Run Now!"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4620
      TabIndex        =   5
      Top             =   2010
      Width           =   1305
   End
   Begin MSComDlg.CommonDialog ctlDialog 
      Left            =   5160
      Top             =   2295
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Sounds (*.wav)|*.wav"
   End
   Begin VB.ListBox lstData 
      Height          =   1425
      Left            =   855
      TabIndex        =   12
      Top             =   810
      Width           =   3270
      Visible         =   0   'False
   End
   Begin VB.CommandButton btnMoveDown 
      Caption         =   "Move Do&wn"
      Enabled         =   0   'False
      Height          =   345
      Left            =   2325
      TabIndex        =   7
      Top             =   2610
      Width           =   2100
   End
   Begin VB.CommandButton btnMoveUp 
      Caption         =   "Move &Up"
      Enabled         =   0   'False
      Height          =   345
      Left            =   165
      TabIndex        =   6
      Top             =   2610
      Width           =   2100
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "&Delete..."
      Enabled         =   0   'False
      Height          =   345
      Left            =   4620
      TabIndex        =   4
      Top             =   1290
      Width           =   1305
   End
   Begin VB.CommandButton btnRename 
      Caption         =   "&Rename..."
      Enabled         =   0   'False
      Height          =   345
      Left            =   4620
      TabIndex        =   3
      Top             =   855
      Width           =   1305
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "&New..."
      Height          =   345
      Left            =   4620
      TabIndex        =   2
      Top             =   420
      Width           =   1305
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3570
      TabIndex        =   10
      Top             =   5265
      Width           =   1125
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4815
      TabIndex        =   11
      Top             =   5265
      Width           =   1125
   End
   Begin HoursAndMinutes.ctlRuleEdit ctlRuleEdit 
      Height          =   1740
      Left            =   165
      TabIndex        =   9
      Top             =   3360
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   3069
   End
   Begin VB.ListBox lstRules 
      Height          =   1950
      IntegralHeight  =   0   'False
      Left            =   165
      TabIndex        =   1
      Top             =   420
      Width           =   4260
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ru&le description (click on an underlined value to edit it):"
      Height          =   195
      Left            =   165
      TabIndex        =   8
      Top             =   3075
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Apply rules in the following order:"
      Height          =   195
      Left            =   165
      TabIndex        =   0
      Top             =   150
      Width           =   2415
   End
End
Attribute VB_Name = "frmRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_blnCancelPressed As Boolean
Public m_strRule As String

Private Sub btnCancel_Click()
    m_blnCancelPressed = True
    Hide
End Sub

Private Sub btnOk_Click()
    Dim strRule As String, strReturn As String, x As Long
    
    ' validate rules and construct rule string
    strRule = ""
    strReturn = ""
    For x = 0 To lstRules.ListCount - 1
        strRule = "1.0" & "|" & lstRules.List(x) & "|" & lstData.List(x)
        If ruleValidate(strRule) = False Then
            lstRules.ListIndex = x
            MsgBox "Unable to parse the rule '" & lstRules.List(x) & "': One or more values has not been set." & vbCrLf & "To set a value, click underlined words at the bottom of the Rules window.", vbExclamation Or vbOKOnly, "Error Parsing Rule"
            Exit Sub
        End If
        strReturn = strReturn & strRule & "*"
    Next
    m_strRule = strReturn
    
    m_blnCancelPressed = False
    Hide
End Sub

Private Sub btnMoveDown_Click()
    Dim strTRule As String, strTData As String, lngIndex As Long
    If lstRules.ListIndex = lstRules.ListCount - 1 Then Exit Sub
    
    lngIndex = lstRules.ListIndex
    
    strTRule = lstRules.List(lngIndex + 1)
    strTData = lstData.List(lngIndex + 1)
    
    lstRules.List(lngIndex + 1) = lstRules.List(lngIndex)
    lstData.List(lngIndex + 1) = lstData.List(lngIndex)
    
    lstRules.List(lngIndex) = strTRule
    lstData.List(lngIndex) = strTData
    
    lstRules.ListIndex = lngIndex + 1
End Sub

Private Sub btnMoveUp_Click()
    Dim strTRule As String, strTData As String, lngIndex As Long
    If lstRules.ListIndex < 1 Then Exit Sub
    
    lngIndex = lstRules.ListIndex
    
    strTRule = lstRules.List(lngIndex - 1)
    strTData = lstData.List(lngIndex - 1)
    
    lstRules.List(lngIndex - 1) = lstRules.List(lngIndex)
    lstData.List(lngIndex - 1) = lstData.List(lngIndex)
    
    lstRules.List(lngIndex) = strTRule
    lstData.List(lngIndex) = strTData
    
    lstRules.ListIndex = lngIndex - 1
End Sub

Private Sub btnNew_Click()
    Load frmRulesAdd
    frmRulesAdd.Show 1, Me
    If frmRulesAdd.m_blnCancelPressed = False Then
        lstRules.AddItem "New Rule " & Now
        lstData.AddItem frmRulesAdd.m_lngWhenType & "|" & frmRulesAdd.m_lngWhatType & "|undefined|undefined|undefined"
        lstRules.ListIndex = lstRules.ListCount - 1
        lstRules_Click
    End If
    Unload frmRulesAdd
End Sub

Private Sub btnDelete_Click()
    If MsgBox("Are you sure you want to remove the rule '" & lstRules.List(lstRules.ListIndex) & "'?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Confirm Rule Delete") = vbYes Then
        lstData.RemoveItem lstRules.ListIndex
        lstRules.RemoveItem lstRules.ListIndex
        lstRules_Click
    End If
End Sub

Private Sub btnRename_Click()
    Dim strInput As String, x As Long
    strInput = getInput(Me, "&New Name:", "Rename", lstRules.List(lstRules.ListIndex))
    If strInput = "" Then Exit Sub
    
    ' check for invalid chars
    If containInvalidChars(strInput) Then
        MsgBox "A rulename cannot contain any of the following characters:" & vbCrLf & "\ / : * ? "" < > |", vbCritical Or vbOKOnly, "Rename"
        Exit Sub
    End If
        
    ' duplicates is ok
        
    ' rename the rule
    lstRules.List(lstRules.ListIndex) = strInput
End Sub

Private Sub btnTest_Click()
    Dim strRule As String
    strRule = "1.0|" & lstRules.List(lstRules.ListIndex) & "|" & lstData.List(lstRules.ListIndex)
    If Not ruleValidate(strRule) Then
        MsgBox "Unable to parse the rule '" & lstRules.List(lstRules.ListIndex) & "': One or more values has not been set." & vbCrLf & "To set a value, click underlined words at the bottom of the Rules window.", vbExclamation Or vbOKOnly, "Error Parsing Rule"
        Exit Sub
    End If
    ruleExecute strRule
End Sub

Private Sub Form_Load()
    m_blnCancelPressed = True
End Sub

Private Sub lstRules_Click()
    Dim arrRule() As String, strText As String
    With lstRules
        If .ListIndex > -1 Then
            btnRename.Enabled = True
            btnDelete.Enabled = True
            btnTest.Enabled = True
            
            If lstRules.ListIndex > 0 Then btnMoveUp.Enabled = True Else btnMoveUp.Enabled = False
            If lstRules.ListIndex + 1 < lstRules.ListCount Then btnMoveDown.Enabled = True Else btnMoveDown.Enabled = False
        Else
            btnRename.Enabled = False
            btnDelete.Enabled = False
            btnTest.Enabled = False
            
            btnMoveUp.Enabled = False
            btnMoveDown.Enabled = False
            ctlRuleEdit.SetText ""
            Exit Sub
        End If
    End With

    arrRule = Split(lstData.List(lstRules.ListIndex), "|")
    Select Case CLng(arrRule(0))
        Case RULE_WHEN_AFTER_X_MINUTES_ONE_DAY
            strText = "Apply this rule after |minutes*" & arrRule(2) & "| minutes of activity in one day" & vbTab
        Case RULE_WHEN_EVERY_X_MINUTES
            strText = "Apply this rule every |minutes*" & arrRule(2) & "| minutes of activity" & vbTab
    End Select
            
    Select Case CLng(arrRule(1))
        Case RULE_WHAT_SHOW_MESSAGE
            strText = strText & "Show the following message '|message*" & arrRule(3) & "|'"
        Case RULE_WHAT_PLAY_SOUND
            strText = strText & "Play the following sound '|sound*" & arrRule(4) & "|'"
        Case RULE_WHAT_SHOW_MESSAGE_PLAY_SOUND
            strText = strText & "Show the following message '|message*" & arrRule(3) & "|'" & vbTab & _
                                "  and Play the following sound '|sound*" & arrRule(4) & "|'"
    End Select
    
    ctlRuleEdit.SetText strText
End Sub

Private Sub ctlRuleEdit_LinkClick(ByVal strLinkID As String, strValue As String, blnSetValue As Boolean)
    Dim arrRule() As String
    
    arrRule = Split(lstData.List(lstRules.ListIndex), "|")
    
    If strLinkID = "minutes" Then
        Load frmRuleMinutes
        With frmRuleMinutes
            strValue = arrRule(2)
            If isInteger(strValue) = False Then strValue = "30"
            .txtHours.Text = CLng(strValue) \ 60
            .txtMinutes.Text = CLng(strValue) - (CLng(.txtHours.Text) * 60)
            .Show 1, Me
            If .m_blnCancelPressed = False Then
                blnSetValue = True
                strValue = CStr((CLng(.txtHours.Text) * 60) + CLng(.txtMinutes.Text))
                arrRule(2) = strValue
            End If
        End With
        Unload frmRuleMinutes
    ElseIf strLinkID = "message" Then
        Load frmRuleMessage
        With frmRuleMessage
            strValue = arrRule(3)
            .txtValue.Text = strValue
            .txtValue.SelStart = 0
            .txtValue.SelLength = Len(.txtValue.Text)
            .Show 1, Me
            If .m_blnCancelPressed = False Then
                blnSetValue = True
                strValue = .txtValue.Text
                arrRule(3) = strValue
            End If
        End With
        Unload frmRuleMessage
    ElseIf strLinkID = "sound" Then
        On Error GoTo lblSkip:
        strValue = arrRule(4)
        ctlDialog.CancelError = True
        ctlDialog.FileName = strValue
        ctlDialog.FLAGS = cdlOFNHideReadOnly Or cdlOFNFileMustExist Or cdlOFNPathMustExist
        ctlDialog.ShowOpen
        blnSetValue = True
        strValue = ctlDialog.FileName
        arrRule(4) = strValue
lblSkip:
    End If
    
    lstData.List(lstRules.ListIndex) = Join(arrRule, "|")
End Sub
