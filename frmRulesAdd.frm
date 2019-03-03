VERSION 5.00
Begin VB.Form frmRulesAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Rule"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4605
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRulesAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   163
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   307
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2070
      TabIndex        =   4
      Top             =   1935
      Width           =   1125
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3315
      TabIndex        =   5
      Top             =   1935
      Width           =   1125
   End
   Begin VB.ComboBox lstWhat 
      Height          =   315
      ItemData        =   "frmRulesAdd.frx":000C
      Left            =   165
      List            =   "frmRulesAdd.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1245
      Width           =   4275
   End
   Begin VB.ComboBox lstWhen 
      Height          =   315
      ItemData        =   "frmRulesAdd.frx":0075
      Left            =   165
      List            =   "frmRulesAdd.frx":0086
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   420
      Width           =   4275
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "W&hat do you want this rule to do?"
      Height          =   195
      Left            =   165
      TabIndex        =   2
      Top             =   975
      Width           =   2445
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&When do you want this rule to start?"
      Height          =   195
      Left            =   165
      TabIndex        =   0
      Top             =   150
      Width           =   2640
   End
End
Attribute VB_Name = "frmRulesAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_blnCancelPressed As Boolean
Public m_lngWhenType As Long
Public m_lngWhatType As Long

Private Sub btnCancel_Click()
    Hide
End Sub

Private Sub btnOk_Click()
    If (lstWhat.ListIndex <= 0) Or (lstWhen.ListIndex <= 0) Then
        MsgBox "Cannot create rule: You must select when to start, and what to do!", vbOKOnly Or vbCritical, "Error Creating Rule"
        Exit Sub
    End If
    
    m_lngWhenType = lstWhen.ItemData(lstWhen.ListIndex)
    m_lngWhatType = lstWhat.ItemData(lstWhat.ListIndex)
    m_blnCancelPressed = False
    Hide
End Sub

Private Sub Form_Load()
    m_blnCancelPressed = True
    lstWhat.ListIndex = 0
    lstWhen.ListIndex = 0
End Sub
