VERSION 5.00
Begin VB.Form frmInput 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hours and Minutes"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInput.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   89
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   341
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2610
      TabIndex        =   2
      Top             =   810
      Width           =   1125
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3825
      TabIndex        =   3
      Top             =   810
      Width           =   1125
   End
   Begin VB.TextBox txtValue 
      Height          =   345
      Left            =   1260
      TabIndex        =   1
      Top             =   165
      Width           =   3675
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      Caption         =   "&New name:"
      Height          =   195
      Left            =   165
      TabIndex        =   0
      Top             =   225
      Width           =   810
   End
End
Attribute VB_Name = "frmInput"
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

