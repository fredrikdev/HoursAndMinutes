VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frmRuleMinutes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Time"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRuleMinutes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   126
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   357
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1410
      TabIndex        =   4
      Top             =   1350
      Width           =   1125
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2700
      TabIndex        =   5
      Top             =   1350
      Width           =   1125
   End
   Begin VB.TextBox txtMinutes 
      Height          =   300
      Left            =   2700
      TabIndex        =   2
      Text            =   "30"
      Top             =   675
      Width           =   570
   End
   Begin VB.TextBox txtHours 
      Height          =   300
      Left            =   900
      TabIndex        =   0
      Text            =   "0"
      Top             =   675
      Width           =   570
   End
   Begin MSComCtl2.UpDown upHours 
      Height          =   300
      Left            =   1470
      TabIndex        =   1
      Top             =   660
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtHours"
      BuddyDispid     =   196612
      OrigLeft        =   94
      OrigTop         =   27
      OrigRight       =   110
      OrigBottom      =   47
      Max             =   100
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown upMinutes 
      Height          =   300
      Left            =   3270
      TabIndex        =   3
      Top             =   660
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   3
      BuddyControl    =   "txtMinutes"
      BuddyDispid     =   196611
      OrigLeft        =   3315
      OrigTop         =   945
      OrigRight       =   3555
      OrigBottom      =   1245
      Max             =   59
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "Adjust the time that will trigger this rule to run:"
      Height          =   240
      Left            =   870
      TabIndex        =   8
      Top             =   150
      Width           =   4815
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Minutes"
      Height          =   195
      Left            =   3630
      TabIndex        =   7
      Top             =   720
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Hours and"
      Height          =   195
      Left            =   1830
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
End
Attribute VB_Name = "frmRuleMinutes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_blnCancelPressed As Boolean

Private Sub btnCancel_Click()
    Hide
End Sub

Private Sub btnOk_Click()
    m_blnCancelPressed = False
    Hide
End Sub

Private Sub Form_Load()
    m_blnCancelPressed = True
End Sub

Private Sub txtHours_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 46) And (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtMinutes_KeyPress(KeyAscii As Integer)
    If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 46) And (KeyAscii <> 8) Then
        KeyAscii = 0
    End If
End Sub
