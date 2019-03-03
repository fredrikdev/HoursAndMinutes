VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hours and Minutes"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Hours and Minutes"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   410
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   604
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer ctlSwitchModeTimer_NewStyle 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5400
      Top             =   60
   End
   Begin HoursAndMinutes.ctlWebbyButton wbTabs 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   35
      Top             =   1605
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   556
      Caption         =   "Preferences"
   End
   Begin MSComDlg.CommonDialog ctlDialog 
      Left            =   7050
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save Backup As"
      Filter          =   "Registration Files|*.reg"
   End
   Begin VB.Timer ctlModeTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6300
      Top             =   60
   End
   Begin VB.Timer ctlSwitchModeTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5850
      Top             =   60
   End
   Begin HoursAndMinutes.ctlTray ctlTray 
      Left            =   6780
      Top             =   240
      _ExtentX        =   423
      _ExtentY        =   423
   End
   Begin VB.Frame ctlTabsPages 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4590
      Index           =   2
      Left            =   2340
      TabIndex        =   25
      Top             =   1200
      Width           =   6540
      Visible         =   0   'False
      Begin VB.ListBox lstStatModes 
         Height          =   2295
         IntegralHeight  =   0   'False
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   46
         Top             =   960
         Width           =   6225
      End
      Begin VB.CommandButton btnStatPlugins 
         Caption         =   "&Plugins >>"
         Height          =   345
         Left            =   4965
         TabIndex        =   14
         Top             =   4020
         Width           =   1365
      End
      Begin VB.CommandButton btnStatCreate 
         Caption         =   "&Create"
         Height          =   345
         Left            =   3495
         TabIndex        =   13
         Top             =   4020
         Width           =   1365
      End
      Begin MSComCtl2.DTPicker dteStatStartDate 
         Height          =   315
         Left            =   810
         TabIndex        =   11
         Top             =   3450
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   16777215
         CalendarTitleForeColor=   0
         CalendarTrailingForeColor=   0
         Format          =   23527425
         UpDown          =   -1  'True
         CurrentDate     =   37101
      End
      Begin MSComCtl2.DTPicker dteStatEndDate 
         Height          =   315
         Left            =   2865
         TabIndex        =   12
         Top             =   3450
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         CalendarBackColor=   16777215
         CalendarForeColor=   0
         CalendarTitleBackColor=   16777215
         CalendarTitleForeColor=   0
         CalendarTrailingForeColor=   0
         Format          =   23527425
         UpDown          =   -1  'True
         CurrentDate     =   37101
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "until"
         Height          =   195
         Left            =   2385
         TabIndex        =   47
         Top             =   3510
         Width           =   300
      End
      Begin VB.Label lblPluginLink 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "plugin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1815
         MouseIcon       =   "frmMain.frx":000C
         MousePointer    =   99  'Custom
         TabIndex        =   32
         ToolTipText     =   "Download Plugins"
         Top             =   405
         Width           =   420
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "&Period:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   3510
         Width           =   510
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Show only time related with the following &task(s):"
         Height          =   240
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   3600
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":0316
         Height          =   480
         Left            =   120
         TabIndex        =   26
         Top             =   210
         Width           =   6225
      End
   End
   Begin VB.Frame ctlTabsPages 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4590
      Index           =   3
      Left            =   2340
      TabIndex        =   28
      Top             =   1200
      Width           =   6540
      Begin VB.CommandButton btnRegister 
         Caption         =   "Register"
         Height          =   345
         Left            =   5220
         TabIndex        =   17
         Top             =   3240
         Width           =   1125
      End
      Begin VB.TextBox txtRegCode 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1845
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   1290
         Width           =   6225
      End
      Begin VB.Label lblMarmaladeMoon 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marmalade Moon"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4800
         MouseIcon       =   "frmMain.frx":03A4
         MousePointer    =   99  'Custom
         TabIndex        =   33
         Top             =   405
         Width           =   1215
      End
      Begin VB.Label lblRegNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note: Your license is personal, sharing the license to others is strictly forbidden!"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   3210
         Width           =   5730
         Visible         =   0   'False
      End
      Begin VB.Label lblRegisteredTo 
         BackStyle       =   0  'Transparent
         Caption         =   "&License:"
         Height          =   240
         Left            =   120
         TabIndex        =   15
         Top             =   1065
         Width           =   3300
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2001 by Port Jackson Computing. All icons from the Marmalade Moon. Used by permission. Artist Kate England."
         Height          =   465
         Left            =   120
         TabIndex        =   30
         Top             =   405
         Width           =   6105
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Port Jackson's...."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         MouseIcon       =   "frmMain.frx":06AE
         MousePointer    =   99  'Custom
         TabIndex        =   29
         ToolTipText     =   "Download Updates"
         Top             =   210
         Width           =   1260
      End
   End
   Begin VB.Frame ctlTabsPages 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4590
      Index           =   0
      Left            =   2340
      TabIndex        =   18
      Top             =   1200
      Width           =   6540
      Begin VB.CheckBox chkAutocomment 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Create comments automatically when switching between tasks."
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   34
         Top             =   3555
         Width           =   5115
      End
      Begin VB.CheckBox chkRemindNewStyle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Us&e new style task-switch reminder."
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   3195
         Value           =   1  'Checked
         Width           =   3870
      End
      Begin VB.TextBox txtMIdle 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2040
         TabIndex        =   1
         Text            =   "5"
         Top             =   540
         Width           =   570
      End
      Begin VB.CheckBox chkMIdle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Stop monitoring after"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   0
         Top             =   585
         Value           =   1  'Checked
         Width           =   1890
      End
      Begin VB.CheckBox chkMSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "A&utosave every"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   105
         TabIndex        =   3
         Top             =   1545
         Value           =   1  'Checked
         Width           =   1890
      End
      Begin VB.TextBox txtMSave 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2040
         TabIndex        =   4
         Text            =   "15"
         Top             =   1500
         Width           =   570
      End
      Begin VB.CheckBox chkModes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Remind me to select task when I go active."
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   2805
         Value           =   1  'Checked
         Width           =   5070
      End
      Begin MSComCtl2.UpDown upMIdle 
         Height          =   300
         Left            =   2610
         TabIndex        =   2
         Top             =   540
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMIdle"
         BuddyDispid     =   196640
         OrigLeft        =   3270
         OrigTop         =   945
         OrigRight       =   3510
         OrigBottom      =   1245
         Max             =   30
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown upMSave 
         Height          =   300
         Left            =   2610
         TabIndex        =   5
         Top             =   1485
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtMSave"
         BuddyDispid     =   196643
         OrigLeft        =   3315
         OrigTop         =   945
         OrigRight       =   3555
         OrigBottom      =   1245
         Max             =   60
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "idle minutes"
         Height          =   195
         Left            =   2970
         TabIndex        =   23
         Top             =   585
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "What would you like Hours and Minutes to do when you're inactive?"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   210
         Width           =   4845
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Do you want Hours and Minutes to autosave your data as a precaution?"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1170
         Width           =   5190
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "minutes"
         Height          =   195
         Left            =   2970
         TabIndex        =   20
         Top             =   1545
         Width           =   555
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":09B8
         Height          =   435
         Left            =   120
         TabIndex        =   19
         Top             =   2295
         Width           =   6240
      End
   End
   Begin VB.Frame ctlTabsPages 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4590
      Index           =   1
      Left            =   2340
      TabIndex        =   24
      Top             =   1200
      Width           =   6540
      Visible         =   0   'False
      Begin MSComctlLib.ListView lstReminders 
         Height          =   3630
         Left            =   120
         TabIndex        =   8
         Top             =   735
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   6403
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "When"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "What"
            Object.Width           =   8062
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Last Run"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Reminders are useful for remembering important dates, or to help you remember different tasks. (Rightclick for options)"
         Height          =   420
         Left            =   120
         TabIndex        =   27
         Top             =   210
         Width           =   6135
      End
   End
   Begin VB.Frame ctlTabsPages 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4590
      Index           =   4
      Left            =   2340
      TabIndex        =   37
      Top             =   1200
      Width           =   6540
      Begin VB.Label lblBackup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Backup Now!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2205
         MouseIcon       =   "frmMain.frx":0A4C
         MousePointer    =   99  'Custom
         TabIndex        =   45
         Top             =   585
         Width           =   930
      End
      Begin VB.Label lblTaskModify 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add, Remove or Modify Tasks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   390
         MouseIcon       =   "frmMain.frx":0D56
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   2385
         Width           =   2145
      End
      Begin VB.Label lblTaskLogEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Edit the Timelog"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   390
         MouseIcon       =   "frmMain.frx":1060
         MousePointer    =   99  'Custom
         TabIndex        =   43
         Top             =   2100
         Width           =   1140
      End
      Begin VB.Label lblTaskAddComent 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comment Task on Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   390
         MouseIcon       =   "frmMain.frx":136A
         MousePointer    =   99  'Custom
         TabIndex        =   42
         Top             =   1815
         Width           =   1665
      End
      Begin VB.Label lblTaskSwitch 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Switch Task"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   390
         MouseIcon       =   "frmMain.frx":1674
         MousePointer    =   99  'Custom
         TabIndex        =   41
         Top             =   1545
         Width           =   840
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":197E
         Height          =   420
         Left            =   390
         TabIndex        =   40
         Top             =   990
         Width           =   6090
      End
      Begin VB.Image Image3 
         Height          =   90
         Left            =   135
         Picture         =   "frmMain.frx":1A05
         Top             =   1035
         Width           =   120
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Secure your data today!"
         Height          =   195
         Left            =   390
         TabIndex        =   39
         Top             =   585
         Width           =   1770
      End
      Begin VB.Image Image2 
         Height          =   90
         Left            =   135
         Picture         =   "frmMain.frx":1A47
         Top             =   645
         Width           =   120
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome to Hours and Minutes!"
         Height          =   465
         Left            =   120
         TabIndex        =   38
         Top             =   210
         Width           =   6270
      End
   End
   Begin VB.Label lblSection 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preferences"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2910
      TabIndex        =   36
      Top             =   780
      Width           =   1455
   End
   Begin VB.Shape shpBorder 
      Height          =   4620
      Left            =   2325
      Top             =   1185
      Width           =   6570
   End
   Begin VB.Menu mnuStatPluginsMenu 
      Caption         =   "Stat Plugins Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuStatPlugins 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuTrayMenu 
      Caption         =   "Tray Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "&Hours and Minutes"
      End
      Begin VB.Menu mnuSwitchMode 
         Caption         =   "&Switch Task"
         Begin VB.Menu mnuModesAdvanced 
            Caption         =   "Show &Advanced"
         End
         Begin VB.Menu mnuLine2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuModes 
            Caption         =   "Mode"
            Index           =   0
         End
      End
      Begin VB.Menu mnuAddComment 
         Caption         =   "Add &Comment..."
      End
      Begin VB.Menu mnuAddReminder 
         Caption         =   "Add &Reminder..."
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuReminderMenu 
      Caption         =   "Reminders Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuReminderAdd 
         Caption         =   "&New..."
      End
      Begin VB.Menu mnuReminderEdit 
         Caption         =   "&Edit..."
      End
      Begin VB.Menu mnuReminderDelete 
         Caption         =   "&Delete..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' public settings:
Public m_blnAllowIdle As Boolean
Public m_lngIdleStart As Long
Public m_blnAllowAutosave As Boolean
Public m_lngAutosaveInterval As Long
Public m_blnRemindOnActive As Boolean
Public m_blnRemindOnActive_NewStyle As Boolean
Public m_blnAutoComment As Boolean

' private variables:
Public m_blnLoadState As Boolean
Private m_blnSwitchVisible As Boolean

' private timing variables
Public m_lngActiveModeID As Long
Public m_dteDate As Date
Public m_lngStartTime As Long
Private m_lngIdleTimeMax As Long

Private m_lngMSeconds As Long
Private m_blnIsIdle As Boolean
Private m_lngLastAutosave As Long
Private m_lngIteration As Long

Private Sub lblBackup_Click()
    Dim dteNow As String, strNow As String, x As Long, y As Long
    Dim lngFile As Long, objReg As clsRegistry
    Dim arrValues() As String, lngValues As Long
    Dim arrSections() As String, lngSections As Long

    On Error GoTo lblError
    If MsgBox("This will backup all Hours and Minutes data to a file." & vbCrLf & "Do you wish to continue?", vbYesNo Or vbQuestion) = vbYes Then
        dteNow = CStr(Now)
        For x = 1 To Len(dteNow)
            If (Mid(dteNow, x, 1) >= "0" And Mid(dteNow, x, 1) <= "9") Or (Mid(dteNow, x, 1) = " ") Then
                strNow = strNow & Mid(dteNow, x, 1)
            End If
        Next
        
        With ctlDialog
            .FileName = "Hours and Minutes Backup " & strNow & ".reg"
            .FLAGS = FileOpenConstants.cdlOFNOverwritePrompt Or FileOpenConstants.cdlOFNPathMustExist
            .ShowSave
            
            On Error GoTo lblErrorWriting
            modeSaveTime m_lngActiveModeID, m_dteDate, m_lngMSeconds
            
            Set objReg = New clsRegistry
            With objReg
                .ClassKey = HKEY_CURRENT_USER
            
                lngFile = FreeFile
                Open ctlDialog.FileName For Output As #lngFile
                Print #lngFile, "REGEDIT4" & vbCrLf
                
                Print #lngFile, "[HKEY_CURRENT_USER\" & m_strRegRoot & "]"
                .SectionKey = m_strRegRoot
                .EnumerateValues arrValues, lngValues
                For x = 1 To lngValues
                    .ValueKey = arrValues(x)
                    Print #lngFile, """" & arrValues(x) & """=""" & regReplace(.Value) & """"
                Next
                Print #lngFile, ""
                
                Print #lngFile, "[HKEY_CURRENT_USER\" & m_strRegRoot & "\Modes]"
                .SectionKey = m_strRegRoot & "\Modes"
                .EnumerateValues arrValues, lngValues
                For x = 1 To lngValues
                    .ValueKey = arrValues(x)
                    Print #lngFile, """" & arrValues(x) & """=""" & regReplace(.Value) & """"
                Next
                Print #lngFile, ""
                
                .EnumerateSections arrSections, lngSections
                For y = 1 To lngSections
                    Print #lngFile, "[HKEY_CURRENT_USER\" & m_strRegRoot & "\Modes\" & arrSections(y) & "]"
                    .SectionKey = m_strRegRoot & "\Modes\" & arrSections(y)
                    .EnumerateValues arrValues, lngValues
                    For x = 1 To lngValues
                        .ValueKey = arrValues(x)
                        Print #lngFile, """" & arrValues(x) & """=""" & regReplace(.Value) & """"
                    Next
                    Print #lngFile, ""
                Next
                
                Close #lngFile
            End With
            Set objReg = Nothing
            MsgBox "The backup has been created successfully!", vbInformation Or vbOKOnly
        End With
    End If
    
    Exit Sub
lblError:
    MsgBox "Warning. No backup has been made.", vbInformation Or vbOKOnly
    Exit Sub
lblErrorWriting:
    MsgBox "There was an error creating the backup:" & vbCrLf & vbCrLf & Err.Description, vbCritical Or vbOKOnly
    Exit Sub
End Sub

Private Sub lblTaskAddComent_Click()
    modeAddComment
End Sub

Private Sub lblTaskLogEdit_Click()
    modeEditLog
End Sub

Private Sub lblTaskModify_Click()
    modeModify
End Sub

Private Sub lblTaskSwitch_Click()
    modeSwitch
End Sub

' this function adds either a switch to or a switch from event to the
' comment string of a certain mode (for today)
Private Sub modeAutoComment(lngModeID As Long, blnSwitchTo As Boolean)
    Dim objReg As clsRegistry
    Set objReg = New clsRegistry
    With objReg
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = m_strRegRoot & "\Modes\" & lngModeID
        .ValueKey = "Comment " & CLng(m_dteDate)
        If blnSwitchTo Then
            .Value = .Value & "Switched to the Task at " & CStr(Time) & "*"
        Else
            .Value = .Value & "Switched from the Task at " & CStr(Time) & "*"
        End If
    End With
    Set objReg = Nothing
End Sub

Private Sub modeAddComment()
    Dim objReg As clsRegistry
    
    Load frmModeComment
    With frmModeComment
        .Show 1, Me
        If .m_blnCancelPressed = False Then
            Set objReg = New clsRegistry
            objReg.ClassKey = HKEY_CURRENT_USER
            objReg.SectionKey = m_strRegRoot & "\Modes\" & m_lngActiveModeID
            objReg.ValueKey = "Comment " & CLng(m_dteDate)
            objReg.Value = objReg.Value & .txtValue.Text & "*"
            Set objReg = Nothing
        End If
    End With
    Unload frmModeComment
End Sub

Private Sub modeEditLog()
    Load frmLogEdit
    With frmLogEdit
        .m_lngFocusModeID = m_lngActiveModeID
        .listModes
        .Show 1, Me
    End With
    Unload frmLogEdit
End Sub

Private Function regReplace(ByVal strString) As String
    strString = Replace(strString, "\", "\\")
    strString = Replace(strString, """", "\""")
    strString = Replace(strString, vbCrLf, vbCr & vbCr & vbLf)
    regReplace = strString
End Function

Private Sub btnRegister_Click()
    Dim strReturn As String, objReg As clsRegistry
    Dim strVal As String
    
    txtRegCode.Enabled = False
    txtRegCode.Text = TrimCRLF(txtRegCode.Text)
    strReturn = regCodeEvaluate(txtRegCode.Text)
    m_lngStartCount = -1
    If strReturn = "OK" Then
        txtRegCode.Locked = True
        txtRegCode.BackColor = vbButtonFace
        btnRegister.Visible = False
        lblRegNote.Visible = True
        If Not m_blnLoadState Then
            MsgBox "Thank you for registering!", vbOKOnly Or vbInformation
            If ctlModeTimer.Enabled = False Then ctlModeTimer.Enabled = True
            Set objReg = New clsRegistry
            With objReg
                .ClassKey = HKEY_CURRENT_USER
                .SectionKey = m_strRegRoot
                .ValueKey = "m_strLicense"
                .Value = txtRegCode.Text
            End With
            Set objReg = Nothing
        End If
    Else
        If Trim(txtRegCode.Text) <> "" Then MsgBox "Error Evaluating License:" & vbCrLf & vbCrLf & strReturn, vbCritical Or vbOKOnly
        txtRegCode.Text = ""
        txtRegCode.Locked = False
        txtRegCode.BackColor = vbWindowBackground
        btnRegister.Visible = True
        lblRegNote.Visible = False
        
        Set objReg = New clsRegistry
        With objReg
            .ClassKey = HKEY_CURRENT_USER
            .SectionKey = m_strRegRoot
            .ValueKey = "m_strLicense"
            .Value = ""
            
            If m_blnLoadState Then
                ' store how many times hours and minutes has been used
                .ClassKey = HKEY_CURRENT_USER
                .SectionKey = "Software\Microsoft\Windows\CurrentVersion\Explorer"
                .ValueKey = "Vodka Lime " & App.Major
                strVal = .Value
                If Not IsNumeric(strVal) Then
                    .Value = "1742"
                Else
                    .Value = CStr(CLng(strVal) + 1)
                End If
                m_lngStartCount = .Value - 1742
            End If
        End With
        Set objReg = Nothing
    End If
    
    txtRegCode.Enabled = True
End Sub

Private Sub btnStatCreate_Click()
    If lstStatModes.ListIndex < 0 Then
        MsgBox "Cannot create report: Please select a task first!", vbCritical Or vbOKOnly, "Error Creating Report"
        Exit Sub
    End If
    
    ' force an autosave
    modeSaveTime m_lngActiveModeID, m_dteDate, m_lngMSeconds
    
    ' load and create
    Load frmStatistics
    frmStatistics.CreateReport lstStatModes.ItemData(lstStatModes.ListIndex), dteStatStartDate.Value, dteStatEndDate.Value
    frmStatistics.Show 1, Me
    Unload frmStatistics
End Sub

Private Sub btnStatPlugins_Click()
    Dim objReg As clsRegistry, x As Long, y As Long, z As Long
    Dim arrTemp() As String
    
    With btnStatPlugins
        ' remove list of plugins
        For x = mnuStatPlugins.Count - 1 To 1 Step -1
            Unload mnuStatPlugins(x)
        Next
        mnuStatPlugins(0).Caption = "<no plugins installed>"
        mnuStatPlugins(0).Enabled = False
        mnuStatPlugins(0).Visible = True
    
        ' check for installed plugins
        Set objReg = New clsRegistry
        With objReg
            .ClassKey = HKEY_CURRENT_USER
            .SectionKey = m_strRegRoot & "\Plugins\Statistics"
            If .EnumerateValues(arrTemp, z) Then
                y = 0
                For x = 1 To z
                    If y <> 0 Then Load mnuStatPlugins(y)
                    With mnuStatPlugins(y)
                        objReg.ValueKey = arrTemp(x)
                        .Enabled = True
                        .Visible = True
                        .Caption = arrTemp(x)
                        .Tag = CStr(objReg.Value)
                    End With
                    y = y + 1
                Next
            End If
        End With
        Set objReg = Nothing
        
        ' show menu
        PopupMenu mnuStatPluginsMenu, , (.Left \ Screen.TwipsPerPixelX) + ctlTabsPages(0).Left, ((.Top + .Height) \ Screen.TwipsPerPixelY) + ctlTabsPages(0).Top
    End With
End Sub

Private Sub ctlModeTimer_Timer()
    Dim lngTickCount As Long, dteDate As Date
    Dim lngIdleTimeTotal As Long, lngIdleTimeToday As Long, strNewTitle As String
    
    ' check if we need to switch mode
    dteDate = Date
    
    ' get tickers
    lngTickCount = GetTickCount
    
    If m_blnAllowIdle Then
        getIdleTime lngTickCount, m_dteDate, lngIdleTimeTotal, lngIdleTimeToday
        If lngIdleTimeTotal >= m_lngIdleStart * 60 * 1000 Then
            If lngIdleTimeToday > m_lngIdleTimeMax Then m_lngIdleTimeMax = lngIdleTimeToday
            If m_blnIsIdle = False Then m_blnIsIdle = True
        Else
            If m_blnIsIdle = True Then
                m_blnIsIdle = False
                m_lngStartTime = m_lngStartTime + m_lngIdleTimeMax
                m_lngIdleTimeMax = 0
                If (m_blnRemindOnActive = True) And (m_blnSwitchVisible = False) Then
                    If m_blnRemindOnActive_NewStyle Then
                        ctlSwitchModeTimer_NewStyle.Enabled = True
                    Else
                        ctlSwitchModeTimer.Enabled = True
                    End If
                End If
            End If
        End If
    Else
        m_blnIsIdle = False
    End If
    
    m_lngMSeconds = lngTickCount - m_lngStartTime - m_lngIdleTimeMax
        
    ' check if an auto-save should be performed
    If (m_blnAllowAutosave = True) Or (m_dteDate <> dteDate) Then
        If ((lngTickCount - m_lngLastAutosave) > (m_lngAutosaveInterval * 60 * 1000)) Or (m_dteDate <> dteDate) Then
            modeSaveTime m_lngActiveModeID, m_dteDate, m_lngMSeconds
            m_lngLastAutosave = lngTickCount
        End If
    End If
    
    ' process any rules & reminders (every 5 iteration to minimize cpu load)
    m_lngIteration = m_lngIteration - 1
    If m_lngIteration <= 0 Then
        ruleProcess m_lngActiveModeID, m_lngMSeconds
        remindersProcess
        m_lngIteration = 5
    End If
    
    ' check if a daybreak is to be performed (if so, an save has just been done)
    If dteDate <> m_dteDate Then
        m_dteDate = dteDate
        m_lngStartTime = lngTickCount - modeGetTime(m_lngActiveModeID, m_dteDate)
        m_lngIdleTimeMax = 0
        getIdleTime lngTickCount, m_dteDate, lngIdleTimeTotal, lngIdleTimeToday
    Else
        ' output some info to our user
        strNewTitle = m_colModes.Item(CStr(m_lngActiveModeID)).m_strName & " [" & FormatMSeconds(m_lngMSeconds) & "]"
        
        ' set tray tooltip
        ctlTray.doTray strNewTitle
        
        ' output some info to our user
        strNewTitle = App.Title & " - " & strNewTitle
        If Caption <> strNewTitle Then Caption = strNewTitle
    End If
End Sub

Private Sub ctlSwitchModeTimer_Timer()
    ctlSwitchModeTimer.Enabled = False
    modeSwitch
End Sub

Private Sub ctlSwitchModeTimer_NewStyle_Timer()
    If (Forms.Count < 2) And (Not m_blnSwitchVisible) Then
        ctlSwitchModeTimer_NewStyle.Enabled = False
        frmModeMessage.showMessage "The active task is " & vbCrLf & m_colModes.Item(CStr(m_lngActiveModeID)).m_strName & " (" & FormatMSeconds(m_lngMSeconds) & ")" & vbCrLf & vbCrLf & "Click here to switch to another task"
    End If
End Sub

Public Sub modeSwitch()
    If Not m_blnSwitchVisible Then
        m_blnSwitchVisible = True
        Load frmModeSwitch
        With frmModeSwitch
            .m_lngFocusModeID = m_lngActiveModeID
            .listModes
            .Show 1, Me
            switchMode .m_lngFocusModeID
        End With
        Unload frmModeSwitch
        m_blnSwitchVisible = False
    End If
End Sub

Private Sub switchMode(lngNewModeID As Long)
    If (lngNewModeID = m_lngActiveModeID) And (m_blnLoadState = False) Then Exit Sub
    
    ' write current mode time to the registry
    If m_blnLoadState = False Then
        If m_lngActiveModeID > -1 Then
            modeSaveTime m_lngActiveModeID, m_dteDate, m_lngMSeconds
            
            If m_blnAutoComment Then
                modeAutoComment m_lngActiveModeID, False
            End If
        End If
    End If
        
    ' load new mode time from the registry
    m_lngActiveModeID = lngNewModeID
    m_dteDate = Date
    m_lngStartTime = GetTickCount - modeGetTime(m_lngActiveModeID, m_dteDate)
    m_lngIteration = 0
    
    If m_blnAutoComment Then
        modeAutoComment m_lngActiveModeID, True
    End If
    
    ' set tray tooltip
    ' ctlTray.doTray App.Title & " - " & m_colModes.Item(CStr(m_lngActiveModeID)).m_strName
End Sub

Private Sub wbTabs_Click(Index As Integer)
    Dim z As Long, x As Long, strPath As String, lngHIcon As Long
    
    For z = 0 To wbTabs.Count - 1
        If z = Index Then
            If z = 2 Then
                ' statistics page (fill modes listbox)
                lstStatModes.Clear
                lstStatModes.AddItem "(all tasks)"
                lstStatModes.ItemData(lstStatModes.ListCount - 1) = -1
                For x = 1 To m_colModes.Count
                    lstStatModes.AddItem m_colModes(x).m_strName
                    lstStatModes.ItemData(lstStatModes.ListCount - 1) = m_colModes(x).m_lngID
                    
                    ' focus active mode
                    If m_colModes.Item(x).m_lngID = m_lngActiveModeID Then
                        lstStatModes.ListIndex = lstStatModes.ListCount - 1
                    End If
                Next
                            
                ' set date span
                dteStatStartDate.Value = Date
                dteStatStartDate.Day = 1
                
                dteStatEndDate.Value = DateAdd("d", -1, DateAdd("M", 1, dteStatStartDate.Value))
            End If
        
            ' load approporiate icon
            strPath = App.Path
            If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
            If z = 0 Then
                strPath = strPath & "Hours and Minutes Icon Preferences.ico" & vbNullChar
                lngHIcon = LoadImage(0, strPath, 1, 32, 32, 16)
            ElseIf z = 1 Then
                strPath = strPath & "Hours and Minutes Icon Reminders.ico" & vbNullChar
                lngHIcon = LoadImage(0, strPath, 1, 32, 32, 16)
            ElseIf z = 2 Then
                strPath = strPath & "Hours and Minutes Icon Statistics.ico" & vbNullChar
                lngHIcon = LoadImage(0, strPath, 1, 32, 32, 16)
            ElseIf z = 3 Or z = 4 Then
                strPath = strPath & "Hours and Minutes Icon Application.ico" & vbNullChar
                lngHIcon = LoadImage(0, strPath, 1, 32, 32, 16)
            End If
            
            ' draw icon
            If lngHIcon > 0 Then
                Me.Cls
                DrawIcon Me.hdc, 155, 45, ByVal lngHIcon
                DestroyIcon lngHIcon
            End If
            
            lblSection.Caption = wbTabs(Index).Caption
            ctlTabsPages(z).Move 156, 80, ctlTabsPages(0).Width, ctlTabsPages(0).Height
            ctlTabsPages(z).Visible = True
            ctlTabsPages(z).ZOrder 0
            
            wbTabs(Index).Selected = True
        Else
            ctlTabsPages(z).Visible = False
            wbTabs(z).Selected = False
        End If
    Next
End Sub

' load settings from registry, write them to the visuals, apply the settings
Private Sub Form_Load()
    Dim objReg As clsRegistry
    Dim x As Long, y As Long
    Dim strTemp As String, blnTemp As Boolean
                            
    ' initialize variables
    m_blnSwitchVisible = False
    
    ' check for previous instance
    If App.PrevInstance = True Then End
    
    ' remove from tasklist
    App.TaskVisible = False
    
    ' initialize about screen
    lblVersion.Caption = App.Title & " v" & App.Major & "." & App.Minor & " build " & App.Revision
    
    ' initialize window
    shpBorder.BorderColor = RGB(70, 86, 120)
    
    ' initialize helpfile path
    strTemp = App.Path
    If Right(strTemp, 1) <> "\" Then strTemp = strTemp & "\"
    strTemp = strTemp & "Hours and Minutes Help.chm"
    App.HelpFile = strTemp
        
    ' initialize tray icon
    strTemp = App.Path
    If Right(strTemp, 1) <> "\" Then strTemp = strTemp & "\"
    strTemp = strTemp & "Hours and Minutes Icon Application.ico"
    ctlTray.m_strIcon = strTemp
    
    ' initialize modes
    modesInitialize
    
    ' set load state
    m_blnLoadState = True
    
    ' load reminders
    remindersGet
    
    ' load settings from registry
    Set objReg = New clsRegistry
    With objReg
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = m_strRegRoot
        
        .ValueKey = "m_blnAllowIdle"
        If CStr(.Value) = "" Then .Value = "1"
        m_blnAllowIdle = IIf(.Value = 1, True, False)
        
        .ValueKey = "m_lngIdleStart"
        If CStr(.Value) = "" Then .Value = "10"
        m_lngIdleStart = .Value
        
        .ValueKey = "m_lngActiveModeID"
        If CStr(.Value) = "" Then .Value = "-1"
        m_lngActiveModeID = .Value
                
        .ValueKey = "m_blnAllowAutosave"
        If CStr(.Value) = "" Then .Value = "1"
        m_blnAllowAutosave = IIf(.Value = 1, True, False)
        
        .ValueKey = "m_lngAutosaveInterval"
        If CStr(.Value) = "" Then .Value = "15"
        m_lngAutosaveInterval = .Value
        
        .ValueKey = "m_blnRemindOnActive"
        If CStr(.Value) = "" Then .Value = "1"
        m_blnRemindOnActive = IIf(.Value = 1, True, False)
        
        .ValueKey = "m_blnRemindOnActive_NewStyle"
        If CStr(.Value) = "" Then .Value = "1"
        m_blnRemindOnActive_NewStyle = IIf(.Value = 1, True, False)
        
        .ValueKey = "m_blnAutoComment"
        If CStr(.Value) = "" Then .Value = "0"
        m_blnAutoComment = IIf(.Value = 1, True, False)
        
        .ValueKey = "m_strLicense"
        txtRegCode.Text = Trim(.Value)
    End With
    Set objReg = Nothing

    ' evaluate registration code
    btnRegister_Click
    blnTemp = True
    If (m_lngStartCount >= 0) And (m_lngStartCount <= 50) Then
        MsgBox "Thank you for trying Hours and Minutes!" & vbCrLf & vbCrLf & _
               "This is a fully functional unregistered version for evaluation use only." & vbCrLf & _
               "The registered version does not display this notice." & vbCrLf & vbCrLf & _
               "You may restart " & (50 - m_lngStartCount) & " more times, before timing will be disabled."
    ElseIf (m_lngStartCount >= 51) Then
        MsgBox "Thank you for trying Hours and Minutes!" & vbCrLf & vbCrLf & _
               "This is a fully functional unregistered version for evaluation use only." & vbCrLf & _
               "The registered version does not display this notice." & vbCrLf & vbCrLf & _
               "Timing has been dissabled. Please register."
        blnTemp = False
    End If

    ' write settings to visual controls
    chkMIdle.Value = IIf(m_blnAllowIdle, 1, 0)
    txtMIdle.Text = m_lngIdleStart
    chkMSave.Value = IIf(m_blnAllowAutosave, 1, 0)
    txtMSave.Text = m_lngAutosaveInterval
    chkModes.Value = IIf(m_blnRemindOnActive, 1, 0)
    chkRemindNewStyle.Value = IIf(m_blnRemindOnActive_NewStyle, 1, 0)
    chkAutocomment.Value = IIf(m_blnAutoComment, 1, 0)
    
    ' initialize tray icon
    ctlTray.doTray ""
    
    ' switch mode
    modeSwitch
    
    ' clear load state
    m_blnLoadState = False
    
    ' initialize form

    
    ' enable timer
    ctlModeTimer.Enabled = blnTemp
                
    ' hide form
    Hide
End Sub

Private Sub doExit()
    Dim objReg As clsRegistry
    Dim objForm As Form
    
    ' disable timer
    ctlModeTimer.Enabled = False
        
    For Each objForm In Forms
        If objForm.Name <> "frmMain" Then
            Unload objForm
        End If
    Next
        
    ' remove tray icon
    ctlTray.Remove
    
    ' save current active mode id to the registry
    Set objReg = New clsRegistry
    With objReg
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = m_strRegRoot
                
        .ValueKey = "m_lngActiveModeID"
        .Value = m_lngActiveModeID
    End With
    Set objReg = Nothing
    
    ' save time to the registry
    modeSaveTime m_lngActiveModeID, m_dteDate, m_lngMSeconds
    
    ' save leave (if autocomment)
    If m_blnAutoComment Then
        modeAutoComment m_lngActiveModeID, False
    End If
        
    ' save reminders to the registry
    remindersSave

    End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim objReg As clsRegistry, z As Long
    
    ' move settings from visuals into variables
    m_blnAllowIdle = IIf(chkMIdle.Value = 1, True, False)
    m_lngIdleStart = txtMIdle.Text
    m_blnAllowAutosave = IIf(chkMSave.Value = 1, True, False)
    m_lngAutosaveInterval = txtMSave.Text
    m_blnRemindOnActive = IIf(chkModes.Value = 1, True, False)
    m_blnRemindOnActive_NewStyle = IIf(chkRemindNewStyle.Value = 1, True, False)
    m_blnAutoComment = IIf(chkAutocomment.Value = 1, True, False)
    
    ' save settings into registry
    If Not m_blnLoadState Then
        Set objReg = New clsRegistry
        With objReg
            .ClassKey = HKEY_CURRENT_USER
            .SectionKey = m_strRegRoot
            
            .ValueKey = "m_blnAllowIdle"
            .Value = IIf(m_blnAllowIdle, 1, 0)
            
            .ValueKey = "m_lngIdleStart"
            .Value = m_lngIdleStart
                                    
            .ValueKey = "m_blnAllowAutosave"
            .Value = IIf(m_blnAllowAutosave, 1, 0)
            
            .ValueKey = "m_lngAutosaveInterval"
            .Value = m_lngAutosaveInterval
            
            .ValueKey = "m_blnRemindOnActive"
            .Value = IIf(m_blnRemindOnActive, 1, 0)
            
            .ValueKey = "m_blnRemindOnActive_NewStyle"
            .Value = IIf(m_blnRemindOnActive_NewStyle, 1, 0)
            
            .ValueKey = "m_blnAutoComment"
            .Value = IIf(m_blnAutoComment, 1, 0)
        End With
        Set objReg = Nothing
    End If
    
    If UnloadMode = vbAppWindows Or UnloadMode = vbAppTaskManager Then
        doExit
    Else
        Hide
        Set Me.Picture = Nothing
        For z = wbTabs.Count - 1 To 1 Step -1
            Unload wbTabs(z)
        Next
        Cancel = 1
    End If
End Sub

Private Sub modeModify()
    Load frmModes
    frmModes.Show 1, Me
    Unload frmModes
End Sub

' enable or disable sub-controls, then enable apply button
Private Sub chkMIdle_Click()
    Dim blnEnabled As Boolean
    blnEnabled = chkMIdle.Value
    
    txtMIdle.Enabled = blnEnabled
    If blnEnabled Then txtMIdle.BackColor = vbWindowBackground Else txtMIdle.BackColor = vbButtonFace
    upMIdle.Enabled = blnEnabled
End Sub

' enable or disable sub-controls, then enable apply button
Private Sub chkMSave_Click()
    Dim blnEnabled As Boolean
    blnEnabled = chkMSave.Value
    
    txtMSave.Enabled = blnEnabled
    If blnEnabled Then txtMSave.BackColor = vbWindowBackground Else txtMSave.BackColor = vbButtonFace
    upMSave.Enabled = blnEnabled
End Sub

Private Sub lblMarmaladeMoon_Click()
    Const SW_SHOWNORMAL = 1
    ShellExecute hwnd, "open", "http://www.marmalademoon.com", vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

Private Sub lblPluginLink_Click()
    Const SW_SHOWNORMAL = 1
    ShellExecute hwnd, "open", "http://www.hoursandminutes.com", vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

Private Sub lblVersion_Click()
    Const SW_SHOWNORMAL = 1
    ShellExecute hwnd, "open", "http://www.hoursandminutes.com", vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

Private Sub mnuAddComment_Click()
    modeAddComment
End Sub

Private Sub mnuModes_Click(Index As Integer)
    switchMode CLng(mnuModes(Index).Tag)
End Sub

Private Sub mnuStatPlugins_Click(Index As Integer)
    Dim objPlugin_Mode() As HAMPluginLib.clsMode
    Dim objPlugin_ModeDay() As HAMPluginLib.clsModeDay
    Dim x As Long, z As Long, objReg As clsRegistry
    Dim lngCurrent As Long, dteCurrent As Date, arrTemp() As String
    Dim objPlugin As Object
    
    If lstStatModes.ListIndex < 0 Then
        MsgBox "Cannot create report: Please select a task first!", vbCritical Or vbOKOnly, "Error Creating Report"
        Exit Sub
    End If
    
    ' force an autosave
    modeSaveTime m_lngActiveModeID, m_dteDate, m_lngMSeconds
    
    On Error GoTo lblError
    
    With mnuStatPlugins(Index)
        Set objPlugin = CreateObject(.Tag)
        If objPlugin.PluginType = "Hours and Minutes Statistics Plugin 1.0" Then
            
            If lstStatModes.ItemData(lstStatModes.ListIndex) = -1 Then
                Set objReg = New clsRegistry
                objReg.ClassKey = HKEY_CURRENT_USER
                
                ReDim objPlugin_Mode(1 To m_colModes.Count)
                ReDim objPlugin_ModeDay(1 To DateDiff("d", dteStatStartDate.Value, dteStatEndDate.Value) + 1)
                
                For x = 1 To m_colModes.Count
                    Set objPlugin_Mode(x) = New HAMPluginLib.clsMode
                    z = 1
                    
                    With objPlugin_Mode(x)
                        objReg.SectionKey = m_strRegRoot & "\Modes\" & m_colModes(x).m_lngID
                        
                        dteCurrent = dteStatStartDate.Value
                        lngCurrent = CLng(dteCurrent)
                        While dteCurrent <= dteStatEndDate.Value
                            Set objPlugin_ModeDay(z) = New HAMPluginLib.clsModeDay
                            With objPlugin_ModeDay(z)
                                .DDate = dteCurrent
                                objReg.ValueKey = "Date " & lngCurrent
                                .Milliseconds = objReg.Value
                                objReg.ValueKey = "Comment " & lngCurrent
                                arrTemp = Split(CStr(objReg.Value), "*")
                                .Comments = arrTemp
                            End With
                        
                            dteCurrent = DateAdd("d", 1, dteCurrent)
                            lngCurrent = CLng(dteCurrent)
                            
                            z = z + 1
                        Wend
                        
                        .modeName = m_colModes(x).m_strName
                        .modeDays = objPlugin_ModeDay
                    End With
                Next
                
                Set objReg = Nothing
                objPlugin.CreateReport objPlugin_Mode
            Else
                Set objReg = New clsRegistry
                objReg.ClassKey = HKEY_CURRENT_USER
                
                ReDim objPlugin_Mode(1 To 1)
                ReDim objPlugin_ModeDay(1 To DateDiff("d", dteStatStartDate.Value, dteStatEndDate.Value) + 1)
                
                Set objPlugin_Mode(1) = New HAMPluginLib.clsMode
                
                x = lstStatModes.ItemData(lstStatModes.ListIndex)
                z = 1
                With objPlugin_Mode(1)
                    objReg.SectionKey = m_strRegRoot & "\Modes\" & m_colModes(CStr(x)).m_lngID
                    
                    dteCurrent = dteStatStartDate.Value
                    lngCurrent = CLng(dteCurrent)
                    While dteCurrent <= dteStatEndDate.Value
                        Set objPlugin_ModeDay(z) = New HAMPluginLib.clsModeDay
                        With objPlugin_ModeDay(z)
                            .DDate = dteCurrent
                            objReg.ValueKey = "Date " & lngCurrent
                            .Milliseconds = objReg.Value
                            objReg.ValueKey = "Comment " & lngCurrent
                            arrTemp = Split(CStr(objReg.Value), "*")
                            .Comments = arrTemp
                        End With
                    
                        dteCurrent = DateAdd("d", 1, dteCurrent)
                        lngCurrent = CLng(dteCurrent)
                        
                        z = z + 1
                    Wend
                    
                    .modeName = m_colModes(CStr(x)).m_strName
                    .modeDays = objPlugin_ModeDay
                End With
                
                Set objReg = Nothing
                objPlugin.CreateReport objPlugin_Mode
            End If
        Else
            MsgBox "This plugin was written for a newer version of Hours and Minutes!", vbCritical Or vbOKOnly
        End If
        Set objPlugin = Nothing
    End With
    
    Exit Sub
lblError:
    MsgBox "There was an error communicating with the Plugin: " & vbCrLf & vbCrLf & Err.Description, vbOKOnly Or vbCritical
End Sub

Private Sub txtMIdle_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 46) And (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtMSave_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 46) And (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

' show main window
Private Sub ctlTray_DblClick(Button As Integer, Shift As Integer, x As Single, y As Single)
    mnuShow_Click
End Sub

' show popup-menu
Private Sub ctlTray_Click(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim z As Long, yz As Long, strAltChars As String
    
    If Button = vbRightButton Then
        If frmMain.Visible = True Then
            mnuShow.Enabled = False
        Else
            mnuShow.Enabled = True
        End If
        
        strAltChars = "hseacr"
        For z = mnuModes.Count - 1 To 1 Step -1
            Unload mnuModes(z)
        Next
        For z = 1 To m_colModes.Count
            If z - 1 = 0 Then
            Else
                Load mnuModes(z - 1)
            End If
            With mnuModes(z - 1)
                .Tag = m_colModes.Item(z).m_lngID
                .Caption = Replace(m_colModes.Item(z).m_strName, "&", "&&")
                
                ' assign first available alt-key to menuitem
                For yz = 1 To Len(.Caption)
                    If InStr(1, strAltChars, Mid$(.Caption, yz, 1), vbTextCompare) < 1 Then
                        strAltChars = strAltChars & Mid$(.Caption, yz, 1)
                        .Caption = Left(.Caption, yz - 1) & "&" & Mid(.Caption, yz)
                        Exit For
                    End If
                Next
               
                ' check the active mode
                .Visible = True
                If m_lngActiveModeID = CLng(.Tag) Then
                    .Checked = True
                Else
                    .Checked = False
                End If
            End With
        Next
        
        PopupMenu mnuTrayMenu, , , , mnuShow
    End If
End Sub

' show main form
Private Sub mnuShow_Click()
    Dim strPath As String
    
    On Error Resume Next
    
    ' load background picture
    strPath = App.Path
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    strPath = strPath & "Hours and Minutes Image Background.gif"
    Me.Picture = LoadPicture(strPath)
    
    ' load buttons
    Load wbTabs(1)
    wbTabs(1).Move 8, 134
    wbTabs(1).Caption = "Reminders"
    wbTabs(1).Visible = True
    
    Load wbTabs(2)
    wbTabs(2).Move 8, 161
    wbTabs(2).Caption = "Statistics"
    wbTabs(2).Visible = True
    
    Load wbTabs(3)
    wbTabs(3).Move 8, 189
    wbTabs(3).Caption = "About"
    wbTabs(3).Visible = True
    
    Load wbTabs(4)
    wbTabs(4).Move 8, 80
    wbTabs(4).Caption = "Welcome"
    wbTabs(4).Visible = True
    
    ' select welcome screen
    wbTabs_Click 4

    Show
End Sub

' exit
Private Sub mnuExit_Click()
    doExit
End Sub

' advanced mode switch
Private Sub mnuModesAdvanced_Click()
    modeSwitch
End Sub

' REMINDERS SECTION -------------------------------------------
' show popupmenu
Private Sub lstReminders_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objItem As ListItem
    If Button <> 2 Then Exit Sub
    
    Set objItem = lstReminders.HitTest(x, y)
    
    If objItem Is Nothing Then
        mnuReminderDelete.Visible = False
        mnuReminderEdit.Visible = False
        mnuReminderAdd.Visible = True
        PopupMenu mnuReminderMenu
    Else
        objItem.Selected = True
        mnuReminderDelete.Tag = objItem.Index
        mnuReminderDelete.Visible = True
        mnuReminderEdit.Tag = objItem.Index
        mnuReminderEdit.Visible = True
        mnuReminderAdd.Visible = True
        PopupMenu mnuReminderMenu
    End If
End Sub

' edit reminder
Private Sub lstReminders_DblClick()
    Dim objItem As ListItem
    
    Set objItem = lstReminders.SelectedItem
    
    If Not objItem Is Nothing Then
        mnuReminderEdit.Tag = objItem.Index
        mnuReminderEdit_Click
    End If
End Sub

' edit reminder
Private Sub mnuReminderEdit_Click()
    ' show edit window
    Load frmReminderEdit
    With frmReminderEdit
        .SetReminder lstReminders.ListItems(CLng(mnuReminderEdit.Tag)).Tag
        .Show 1, Me
        If .m_blnCancelPressed = False Then
            With lstReminders.ListItems(CLng(mnuReminderEdit.Tag))
                .SubItems(1) = frmReminderEdit.txtName.Text
                .SubItems(2) = ""
                .Text = FormatDateTime(frmReminderEdit.dteRemindDate.Value, vbShortDate) & " " & FormatDateTime(frmReminderEdit.dteRemindTime.Value, vbShortTime)
                .Tag = frmReminderEdit.GetReminder
            End With
            remindersSave
        End If
    End With
    Unload frmReminderEdit
End Sub

' delete reminder
Private Sub mnuReminderDelete_Click()
    ' delete reminder with index mnuReminderDelete.Tag
    If MsgBox("Are you sure that you wish to delete the selected reminder?", vbYesNo Or vbQuestion) = vbYes Then
        lstReminders.ListItems.Remove CLng(mnuReminderDelete.Tag)
        remindersSave
    End If
End Sub

' add new reminder
Private Sub mnuReminderAdd_Click()
    ' show window
    Load frmReminderEdit
    With frmReminderEdit
        .Show 1, Me
        If .m_blnCancelPressed = False Then
            With lstReminders.ListItems.Add(, , FormatDateTime(.dteRemindDate.Value, vbShortDate) & " " & FormatDateTime(.dteRemindTime.Value, vbShortTime))
                .Tag = frmReminderEdit.GetReminder
                .SubItems(1) = frmReminderEdit.txtName.Text
                .SubItems(2) = ""
            End With
            remindersSave
        End If
    End With
    Unload frmReminderEdit
End Sub

Private Sub mnuAddReminder_Click()
    ' show window
    mnuReminderAdd_Click
End Sub

