VERSION 5.00
Begin VB.Form frmRuleMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Message"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRuleMessage.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   147
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   339
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2565
      TabIndex        =   2
      Top             =   1665
      Width           =   1125
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3795
      TabIndex        =   3
      Top             =   1665
      Width           =   1125
   End
   Begin VB.TextBox txtValue 
      Height          =   1065
      Left            =   165
      TabIndex        =   1
      Top             =   405
      Width           =   4740
   End
   Begin VB.Label Label3 
      Caption         =   "&Adjust the message that will appear when this rule run:"
      Height          =   240
      Left            =   165
      TabIndex        =   0
      Top             =   150
      Width           =   4815
   End
End
Attribute VB_Name = "frmRuleMessage"
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
    If containInvalidChars(txtValue.Text) Then
        MsgBox "A message cannot contain any of the following characters:" & vbCrLf & "\ / : * ? "" < > |", vbCritical Or vbOKOnly, "Message"
        Exit Sub
    End If
    m_blnCancelPressed = False
    Hide
End Sub

Private Sub Form_Load()
    m_blnCancelPressed = True
End Sub
