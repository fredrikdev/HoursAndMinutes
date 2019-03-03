VERSION 5.00
Begin VB.Form frmModeSwitch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hours and Minutes - Switch Task"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
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
   Icon            =   "frmModeSwitch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnExit 
      Caption         =   "Exit"
      Height          =   345
      Left            =   120
      TabIndex        =   4
      Top             =   1890
      Width           =   1125
      Visible         =   0   'False
   End
   Begin VB.CommandButton btnAddMode 
      Caption         =   "&New Task..."
      Height          =   345
      Left            =   3510
      TabIndex        =   3
      Top             =   1890
      Width           =   1125
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2235
      TabIndex        =   2
      Top             =   1890
      Width           =   1125
   End
   Begin VB.ComboBox lstModes 
      Height          =   315
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1110
      Width           =   4485
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "&Switch to this task:"
      Height          =   195
      Left            =   165
      TabIndex        =   0
      Top             =   855
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Please select the task that you want to log time to. You may setup a new task by clicking the 'New Task' button."
      Height          =   630
      Left            =   165
      TabIndex        =   5
      Top             =   150
      Width           =   4440
   End
End
Attribute VB_Name = "frmModeSwitch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_lngFocusModeID As Long

Private Sub btnAddMode_Click()
    SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    Load frmModes
    frmModes.Show 1, Me
    Unload frmModes
    listModes
    If lstModes.ListIndex < 0 Then
        lstModes.ListIndex = lstModes.ListCount - 1
    End If
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Sub

Private Sub btnExit_Click()
    frmMain.ctlTray.Remove
    End
End Sub

Private Sub btnOk_Click()
    If lstModes.ListIndex < 0 Then
        MsgBox "Error Selecting Mode: Please select a task to use!", vbExclamation Or vbOKOnly
        Exit Sub
    End If
    m_lngFocusModeID = lstModes.ItemData(lstModes.ListIndex)
    Hide
End Sub

Private Sub Form_Load()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    m_lngFocusModeID = -1
    If frmMain.m_blnLoadState Then
        btnExit.Visible = True
    Else
        btnExit.Visible = False
    End If
End Sub

Public Sub listModes()
    Dim x As Long
    lstModes.Clear
    
    For x = 1 To m_colModes.Count
        With m_colModes.Item(x)
            lstModes.AddItem .m_strName
            lstModes.ItemData(lstModes.ListCount - 1) = .m_lngID
            If m_lngFocusModeID = .m_lngID Then
                lstModes.ListIndex = lstModes.ListCount - 1
            End If
        End With
    Next
End Sub
