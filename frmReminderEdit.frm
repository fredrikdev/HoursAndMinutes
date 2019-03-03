VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmReminderEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reminder"
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReminderEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   342
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   437
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame ctlWizard 
      BorderStyle     =   0  'None
      Height          =   4410
      Index           =   0
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   6555
      Begin VB.CheckBox chkRemindOptions 
         Caption         =   "Remind &me every month at the specified day && time."
         Height          =   240
         Index           =   4
         Left            =   375
         TabIndex        =   8
         Top             =   3450
         Width           =   4125
      End
      Begin VB.CheckBox chkRemindOptions 
         Caption         =   "Remin&d me every year at the specified date && time"
         Height          =   240
         Index           =   3
         Left            =   375
         TabIndex        =   7
         Top             =   3120
         Width           =   4170
      End
      Begin VB.CheckBox chkRemindOptions 
         Caption         =   "R&emind me at the specified date && time"
         Height          =   240
         Index           =   0
         Left            =   375
         TabIndex        =   4
         Top             =   2130
         Value           =   1  'Checked
         Width           =   3630
      End
      Begin VB.CheckBox chkRemindOptions 
         Caption         =   "Rem&ind me every 'wednesday' at the specified time"
         Height          =   240
         Index           =   2
         Left            =   375
         TabIndex        =   6
         Top             =   2790
         Width           =   5010
      End
      Begin VB.CheckBox chkRemindOptions 
         Caption         =   "Re&mind me everyday at the specified time (regardless of date)"
         Height          =   240
         Index           =   1
         Left            =   375
         TabIndex        =   5
         Top             =   2460
         Width           =   5010
      End
      Begin MSComCtl2.DTPicker dteRemindDate 
         Height          =   330
         Left            =   2595
         TabIndex        =   3
         Top             =   1515
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   19726337
         UpDown          =   -1  'True
         CurrentDate     =   37102
      End
      Begin MSComCtl2.DTPicker dteRemindTime 
         Height          =   330
         Left            =   2595
         TabIndex        =   1
         Top             =   1065
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   19726338
         UpDown          =   -1  'True
         CurrentDate     =   37102
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "At this specific &date"
         Height          =   195
         Left            =   375
         TabIndex        =   2
         Top             =   1575
         Width           =   1410
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "I would like to be &reminded at"
         Height          =   195
         Left            =   375
         TabIndex        =   0
         Top             =   1125
         Width           =   2115
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "When do you want to be reminded about this event? Please select from the options below."
         Height          =   435
         Left            =   375
         TabIndex        =   21
         Top             =   345
         Width           =   5940
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "When"
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
         TabIndex        =   20
         Top             =   135
         Width           =   480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   6555
         Y1              =   885
         Y2              =   885
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
   Begin VB.Frame ctlWizard 
      BorderStyle     =   0  'None
      Height          =   4410
      Index           =   1
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   6555
      Begin VB.CheckBox chkSound 
         Caption         =   "&Play the following sound:"
         Height          =   285
         Left            =   375
         TabIndex        =   11
         Top             =   2820
         Width           =   3930
      End
      Begin VB.CheckBox chkMessage 
         Caption         =   "&Show this message:"
         Height          =   195
         Left            =   390
         TabIndex        =   9
         Top             =   1140
         Value           =   1  'Checked
         Width           =   2640
      End
      Begin VB.CommandButton btnSoundBrowse 
         Caption         =   "B&rowse..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   360
         TabIndex        =   13
         Top             =   3615
         Width           =   1125
      End
      Begin VB.TextBox txtSound 
         BackColor       =   &H8000000F&
         Height          =   300
         Left            =   375
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   3150
         Width           =   5850
      End
      Begin VB.TextBox txtMessage 
         Height          =   975
         Left            =   375
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   1455
         Width           =   5850
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "What"
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
         TabIndex        =   24
         Top             =   135
         Width           =   450
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   6555
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "What would you like to be reminded of? Optionally, you may even play a sound when the reminder is shown."
         Height          =   435
         Left            =   375
         TabIndex        =   23
         Top             =   345
         Width           =   5940
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   930
         Left            =   0
         Top             =   0
         Width           =   6570
      End
   End
   Begin VB.Frame ctlWizard 
      BorderStyle     =   0  'None
      Height          =   4410
      Index           =   2
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   6555
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   375
         TabIndex        =   15
         Top             =   1455
         Width           =   5850
      End
      Begin VB.Label Label9 
         Caption         =   "N&ame this reminder:"
         Height          =   225
         Left            =   375
         TabIndex        =   14
         Top             =   1140
         Width           =   4500
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Finally, please enter a descriptive name for the reminder."
         Height          =   435
         Left            =   375
         TabIndex        =   27
         Top             =   345
         Width           =   5940
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   6555
         Y1              =   885
         Y2              =   885
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         TabIndex        =   26
         Top             =   135
         Width           =   480
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   930
         Left            =   0
         Top             =   0
         Width           =   6570
      End
   End
   Begin MSComDlg.CommonDialog ctlDialog 
      Left            =   5940
      Top             =   3735
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Sounds (*.wav)|*.wav"
   End
   Begin VB.CommandButton btnBack 
      Caption         =   "<< &Back"
      Height          =   345
      Left            =   2865
      TabIndex        =   16
      Top             =   4635
      Width           =   1125
   End
   Begin VB.CommandButton btnNext 
      Caption         =   "&Next >>"
      Height          =   345
      Left            =   3990
      TabIndex        =   17
      Top             =   4635
      Width           =   1125
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   5280
      TabIndex        =   18
      Top             =   4635
      Width           =   1125
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   437
      Y1              =   295
      Y2              =   295
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   437
      Y1              =   294
      Y2              =   294
   End
End
Attribute VB_Name = "frmReminderEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lngVisiblePage As Long
Public m_blnCancelPressed As Boolean

' sets a reminder for editing
Public Sub SetReminder(strReminder As String)
    ' set the reminder
    ' version|name|when-mask|when-date|when-time|what-mask|what-message|what-sound
    Dim arrReminder() As String
    
    arrReminder = Split(strReminder, "|")
    If UBound(arrReminder) < 7 Then Exit Sub
    If arrReminder(0) = "1.0" Then
        ' set the name parameter
        txtName.Text = arrReminder(1)
    
        ' parse the when-mask parameter
        If (CLng(arrReminder(2)) And enmReminderWhen.REMIND_AT_SPECIFIED_DATE_AND_TIME) = enmReminderWhen.REMIND_AT_SPECIFIED_DATE_AND_TIME Then
            chkRemindOptions(0).Value = vbChecked
        Else
            chkRemindOptions(0).Value = vbUnchecked
        End If
        
        If (CLng(arrReminder(2)) And enmReminderWhen.REMIND_EVERYDAY_AT_SPECIFIED_TIME) = enmReminderWhen.REMIND_EVERYDAY_AT_SPECIFIED_TIME Then
            chkRemindOptions(1).Value = vbChecked
        Else
            chkRemindOptions(1).Value = vbUnchecked
        End If
        
        If (CLng(arrReminder(2)) And enmReminderWhen.REMIND_EVERY_WEEKDAY_AT_SPECIFIED_TIME) = enmReminderWhen.REMIND_EVERY_WEEKDAY_AT_SPECIFIED_TIME Then
            chkRemindOptions(2).Value = vbChecked
        Else
            chkRemindOptions(2).Value = vbUnchecked
        End If
        
        If (CLng(arrReminder(2)) And enmReminderWhen.REMIND_EVERY_YEAR_AT_SPECIFIED_DATE_AND_TIME) = enmReminderWhen.REMIND_EVERY_YEAR_AT_SPECIFIED_DATE_AND_TIME Then
            chkRemindOptions(3).Value = vbChecked
        Else
            chkRemindOptions(3).Value = vbUnchecked
        End If
        
        If (CLng(arrReminder(2)) And enmReminderWhen.REMIND_EVERY_MONTH_AT_SPECIFIED_DATE_AND_TIME) = enmReminderWhen.REMIND_EVERY_MONTH_AT_SPECIFIED_DATE_AND_TIME Then
            chkRemindOptions(4).Value = vbChecked
        Else
            chkRemindOptions(4).Value = vbUnchecked
        End If
        
        ' set the when-parameters
        dteRemindDate.Value = arrReminder(3)
        dteRemindTime.Value = arrReminder(4)
        
        ' parse the what-mask parameter
        If (CLng(arrReminder(5)) And enmReminderWhat.SHOW_MESSAGE) = enmReminderWhat.SHOW_MESSAGE Then
            chkMessage.Value = vbChecked
        Else
            chkMessage.Value = vbUnchecked
        End If
        chkMessage_Click
        
        If (CLng(arrReminder(5)) And enmReminderWhat.PLAY_SOUND) = enmReminderWhat.PLAY_SOUND Then
            chkSound.Value = vbChecked
        Else
            chkSound.Value = vbUnchecked
        End If
        chkSound_Click
        
        ' set the what-parameters
        txtMessage.Text = arrReminder(6)
        txtSound.Text = arrReminder(7)
        
        dteRemindDate_Change
    End If
End Sub

' gets a reminder after editing
Public Function GetReminder() As String
    Dim strReminder As String, lngTemp As Long
    
    strReminder = "1.0|" & txtName.Text & "|"
    
    ' build when-mask parameter
    lngTemp = 0
    If chkRemindOptions(0).Value = vbChecked Then lngTemp = lngTemp Or enmReminderWhen.REMIND_AT_SPECIFIED_DATE_AND_TIME
    If chkRemindOptions(1).Value = vbChecked Then lngTemp = lngTemp Or enmReminderWhen.REMIND_EVERYDAY_AT_SPECIFIED_TIME
    If chkRemindOptions(2).Value = vbChecked Then lngTemp = lngTemp Or enmReminderWhen.REMIND_EVERY_WEEKDAY_AT_SPECIFIED_TIME
    If chkRemindOptions(3).Value = vbChecked Then lngTemp = lngTemp Or enmReminderWhen.REMIND_EVERY_YEAR_AT_SPECIFIED_DATE_AND_TIME
    If chkRemindOptions(4).Value = vbChecked Then lngTemp = lngTemp Or enmReminderWhen.REMIND_EVERY_MONTH_AT_SPECIFIED_DATE_AND_TIME
    strReminder = strReminder & lngTemp & "|"
    
    ' set when-date & when-time parameter
    strReminder = strReminder & FormatDateTime(dteRemindDate.Value, vbShortDate) & "|" & FormatDateTime(dteRemindTime.Value, vbShortTime) & "|"
    
    ' build what-mask parameter
    lngTemp = 0
    If chkMessage.Value = vbChecked Then lngTemp = lngTemp Or enmReminderWhat.SHOW_MESSAGE
    If chkSound.Value = vbChecked Then lngTemp = lngTemp Or enmReminderWhat.PLAY_SOUND
    strReminder = strReminder & lngTemp & "|"
    
    ' set message & sound parameter
    strReminder = strReminder & txtMessage.Text & "|" & txtSound.Text & "|"
        
    GetReminder = strReminder
End Function

' window functions (for wizards and interactivity) ------------------------------------------------
Private Sub btnCancel_Click()
    Hide
End Sub

Private Sub btnNext_Click()
    If btnNext.Caption = "&Next >>" Then
        swWizardPage m_lngVisiblePage + 1
    Else
        If Trim(txtName.Text) = "" Then
            MsgBox "Unable to create reminder: Please enter a name for the reminder.", vbExclamation Or vbOKOnly, "Error Creating Reminder"
            Exit Sub
        ElseIf containInvalidChars(txtName.Text) Then
            MsgBox "A reminder-name cannot contain any of the following characters:" & vbCrLf & "\ / : * ? "" < > |", vbCritical Or vbOKOnly
            Exit Sub
        End If
        m_blnCancelPressed = False
        Hide
    End If
End Sub

Private Sub btnBack_Click()
    swWizardPage m_lngVisiblePage - 1
End Sub

Private Sub btnSoundBrowse_Click()
    On Error GoTo lblError
    ctlDialog.FLAGS = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
    ctlDialog.FileName = txtSound.Text
    ctlDialog.ShowOpen
    txtSound.Text = ctlDialog.FileName
lblError:
End Sub

Private Sub chkMessage_Click()
    If chkMessage.Value = vbChecked Then
        txtMessage.BackColor = vbWindowBackground
        txtMessage.Enabled = True
    Else
        txtMessage.BackColor = vbButtonFace
        txtMessage.Enabled = False
    End If
End Sub

Private Sub chkSound_Click()
    If chkSound.Value = vbChecked Then
        txtSound.BackColor = vbWindowBackground
        btnSoundBrowse.Enabled = True
    Else
        txtSound.BackColor = vbButtonFace
        btnSoundBrowse.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    m_blnCancelPressed = True
    dteRemindTime.Value = DateAdd("n", 30, Now)
    dteRemindDate.Value = dteRemindTime.Value
    dteRemindDate_Change
    swWizardPage 0
End Sub

Private Sub dteRemindDate_Change()
    chkRemindOptions(2).Caption = "Remind me every '" & DayName(dteRemindDate.Value) & "' at the specified time"
End Sub

Private Sub swWizardPage(lngNum As Long)
    Dim X As Long, blnTemp As Boolean
    
    If lngNum = 1 Then
        blnTemp = False
        For X = 0 To chkRemindOptions.Count - 1
            If chkRemindOptions(X).Value = vbChecked Then
                blnTemp = True
                Exit For
            End If
        Next
        If blnTemp = False Then
            MsgBox "You need to select at least one of the checkboxes!", vbCritical Or vbOKOnly
            Exit Sub
        End If
    ElseIf lngNum = 2 Then
        If containInvalidChars(txtMessage.Text) Then
            MsgBox "A userdefined message cannot contain any of the following characters:" & vbCrLf & "\ / : * ? "" < > |", vbCritical Or vbOKOnly
            Exit Sub
        End If
        If (Trim(txtMessage.Text) = "") And (chkMessage.Value = vbChecked) Then
            MsgBox "You need to write a message before clicking next!", vbCritical Or vbOKOnly
            Exit Sub
        End If
        If (chkMessage.Value = vbUnchecked) And (chkSound.Value = vbUnchecked) Then
            MsgBox "You need to select at least one of the checkboxes!", vbCritical Or vbOKOnly
            Exit Sub
        End If
        If (txtSound.Text = "") And (chkSound.Value = vbChecked) Then
            MsgBox "You need to select a sound before clicking next!", vbCritical Or vbOKOnly
            Exit Sub
        End If
    End If
    
    If lngNum + 1 >= ctlWizard.Count Then
        btnNext.Caption = "&Finish"
    Else
        btnNext.Caption = "&Next >>"
    End If
    If lngNum = 0 Then
        btnBack.Enabled = False
    Else
        btnBack.Enabled = True
    End If
    
    For X = 0 To ctlWizard.Count - 1
        If lngNum = X Then
            ctlWizard(X).Visible = True
        Else
            ctlWizard(X).Visible = False
        End If
    Next
    ctlWizard(lngNum).ZOrder 0
    
    m_lngVisiblePage = lngNum
    
    On Error Resume Next
    If lngNum = 0 Then
        If ctlWizard(lngNum).Visible Then dteRemindTime.SetFocus
    ElseIf lngNum = 1 Then
        If ctlWizard(lngNum).Visible Then txtMessage.SetFocus
    ElseIf lngNum = 2 Then
        If ctlWizard(lngNum).Visible Then txtName.SetFocus
    End If
End Sub
