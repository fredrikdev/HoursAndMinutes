VERSION 5.00
Begin VB.Form frmModes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modify Tasks"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmModes.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   384
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstRules 
      Height          =   1230
      Left            =   165
      TabIndex        =   8
      Top             =   2220
      Width           =   2880
      Visible         =   0   'False
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3210
      TabIndex        =   5
      Top             =   3075
      Width           =   1125
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4485
      TabIndex        =   6
      Top             =   3075
      Width           =   1125
   End
   Begin VB.CommandButton btnRules 
      Caption         =   "R&ules..."
      Enabled         =   0   'False
      Height          =   345
      Left            =   4485
      TabIndex        =   4
      Top             =   2100
      Width           =   1125
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "&Delete..."
      Enabled         =   0   'False
      Height          =   345
      Left            =   4485
      TabIndex        =   3
      Top             =   1635
      Width           =   1125
   End
   Begin VB.CommandButton btnRename 
      Caption         =   "&Rename..."
      Enabled         =   0   'False
      Height          =   345
      Left            =   4485
      TabIndex        =   2
      Top             =   1170
      Width           =   1125
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "&New..."
      Height          =   345
      Left            =   4485
      TabIndex        =   1
      Top             =   705
      Width           =   1125
   End
   Begin VB.ListBox lstModes 
      Height          =   1755
      IntegralHeight  =   0   'False
      Left            =   150
      TabIndex        =   0
      Top             =   705
      Width           =   4170
   End
   Begin VB.Label lblInfo 
      Caption         =   "Tasks are used for specifying exactly what you are doing at a certain hour, ie. Working, Playing Games etc."
      Height          =   405
      Left            =   165
      TabIndex        =   7
      Top             =   150
      Width           =   5475
   End
End
Attribute VB_Name = "frmModes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_strRemovedIDs As String

Private Sub btnOk_Click()
    Dim arrRemovedIDs() As String, x As Long
    
    ' delete deleted modes
    arrRemovedIDs = Split(m_strRemovedIDs, ",")
    If IsArray(arrRemovedIDs) Then
        For x = LBound(arrRemovedIDs) To UBound(arrRemovedIDs)
            If isInteger(arrRemovedIDs(x)) = True And arrRemovedIDs(x) <> "-1" Then
                modeDelete CLng(arrRemovedIDs(x))
            End If
        Next
    End If
    
    ' update visible modes
    For x = 0 To lstModes.ListCount - 1
        modeUpdate lstModes.List(x), lstRules.List(x), lstModes.ItemData(x)
    Next
    
    ' write modes to registry
    modesStore
    
    Hide
End Sub

Private Sub btnCancel_Click()
    Hide
End Sub

Private Sub btnNew_Click()
    Dim strInput As String, x As Long
    strInput = getInput(Me, "&New Task:", "New", "")
    If strInput = "" Then Exit Sub
    
    ' check for invalid chars
    If containInvalidChars(strInput) Then
        MsgBox "A taskname cannot contain any of the following characters:" & vbCrLf & "\ / : * ? "" < > |", vbCritical Or vbOKOnly, "New"
        Exit Sub
    End If
    
    ' check for duplicates
    For x = 0 To lstModes.ListCount - 1
        If UCase(lstModes.List(x)) = UCase(strInput) Then
            MsgBox "Cannot create " & strInput & ": A task with the name you specified already exists. Specify a different taskname.", vbCritical Or vbOKOnly, "Error Creating Task"
            Exit Sub
        End If
    Next
    
    ' add the mode
    lstModes.AddItem strInput
    lstModes.ItemData(lstModes.ListCount - 1) = -1
    lstRules.AddItem ""
    
    ' focus the mode
    For x = 0 To lstModes.ListCount - 1
        If UCase(lstModes.List(x)) = UCase(strInput) Then
            lstModes.ListIndex = x
            Exit Sub
        End If
    Next
End Sub

Private Sub btnRename_Click()
    Dim strInput As String, x As Long
    strInput = getInput(Me, "&New Name:", "Rename", lstModes.List(lstModes.ListIndex))
    If strInput = "" Then Exit Sub
    
    ' check for invalid chars
    If containInvalidChars(strInput) Then
        MsgBox "A taskname cannot contain any of the following characters:" & vbCrLf & "\ / : * ? "" < > |", vbCritical Or vbOKOnly, "Rename"
        Exit Sub
    End If
    
    ' check for duplicates
    For x = 0 To lstModes.ListCount - 1
        If (x <> lstModes.ListIndex) And (UCase(lstModes.List(x)) = UCase(strInput)) Then
            MsgBox "Cannot rename " & lstModes.List(lstModes.ListIndex) & ": A task with the name you specified already exists. Specify a different taskname.", vbCritical Or vbOKOnly, "Error Renaming Task"
            Exit Sub
        End If
    Next
    
    ' rename the mode
    lstModes.List(lstModes.ListIndex) = strInput
End Sub

Private Sub btnDelete_Click()
    If lstModes.ItemData(lstModes.ListIndex) = frmMain.m_lngActiveModeID Then
        MsgBox "Cannot delete " & lstModes.List(lstModes.ListIndex) & ": The task is currently in use.", vbCritical Or vbOKOnly, "Error Deleting Task"
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to remove the task '" & lstModes.List(lstModes.ListIndex) & "' and all its associated data?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Confirm Task Delete") = vbYes Then
        m_strRemovedIDs = m_strRemovedIDs & lstModes.ItemData(lstModes.ListIndex) & ","
        lstRules.RemoveItem lstModes.ListIndex
        lstModes.RemoveItem lstModes.ListIndex
    End If
    lstModes_Click
End Sub

Private Sub btnRules_Click()
    Dim x As Long, arrRules() As String
    
    Load frmRules
    
    With frmRules
        ' upload rule data to rule edit window
        .lstData.Clear
        .lstRules.Clear
        arrRules = Split(lstRules.List(lstModes.ListIndex), "*")
        For x = LBound(arrRules) To UBound(arrRules)
            If ruleValidate(arrRules(x)) = True Then
                .lstRules.AddItem ruleGetName(arrRules(x))
                .lstData.AddItem ruleGetData(arrRules(x))
            End If
        Next
        .Show 1, Me
        
        ' download rule data from the rule edit window
        If Not .m_blnCancelPressed Then
            lstRules.List(lstModes.ListIndex) = .m_strRule
        End If
    End With
        
    Unload frmRules
End Sub

Private Sub Form_Load()
    Dim x As Long
    
    m_strRemovedIDs = ""
    For x = 1 To m_colModes.Count
        lstModes.AddItem m_colModes.Item(x).m_strName
        lstModes.ItemData(lstModes.ListCount - 1) = m_colModes.Item(x).m_lngID
        lstRules.AddItem m_colModes.Item(x).m_strRule
    Next
End Sub

Private Sub lstModes_Click()
    With lstModes
        If .ListIndex > -1 Then
            btnRename.Enabled = True
            btnDelete.Enabled = True
            btnRules.Enabled = True
        Else
            btnRename.Enabled = False
            btnDelete.Enabled = False
            btnRules.Enabled = False
        End If
    End With
End Sub
