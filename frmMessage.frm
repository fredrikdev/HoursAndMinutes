VERSION 5.00
Begin VB.Form frmMessage 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hours and Minutes"
   ClientHeight    =   2550
   ClientLeft      =   150
   ClientTop       =   435
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
   HasDC           =   0   'False
   Icon            =   "frmMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   170
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   357
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnNext 
      Caption         =   "&Next >>"
      Height          =   345
      Left            =   4050
      TabIndex        =   0
      Top             =   2040
      Width           =   1125
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   345
      Left            =   2805
      TabIndex        =   1
      Top             =   2040
      Width           =   1125
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   1725
      Index           =   1
      Left            =   855
      TabIndex        =   2
      Top             =   120
      Width           =   4320
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lngVisibleMessage As Long
Private m_lngHIconSmall As Long, m_lngHIconLarge As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Sub Form_Load()
    Dim strIcon As String
    
    ' set icon
    strIcon = App.Path
    If Right(strIcon, 1) <> "\" Then strIcon = strIcon & "\"
    strIcon = strIcon & "Hours and Minutes Icon Reminders.ico"
    m_lngHIconSmall = LoadImage(0, strIcon & vbNullChar, 1, 16, 16, 16)
    m_lngHIconLarge = LoadImage(0, strIcon & vbNullChar, 1, 32, 32, 16)
    If m_lngHIconSmall > 0 Then SendMessage hWnd, &H80, 0, ByVal m_lngHIconSmall
    If m_lngHIconLarge > 0 Then DrawIcon hdc, 11, 10, ByVal m_lngHIconLarge

    ' set windows pos
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE
    btnOk.Left = 270
    btnNext.Visible = False
    m_lngVisibleMessage = -1
End Sub

Private Sub btnOk_Click()
    Dim X As Long
    For X = lblMessage.Count To 2 Step -1
        Unload lblMessage(X)
    Next
    Unload Me
End Sub

Private Sub btnNext_Click()
    swMessage m_lngVisibleMessage + 1
    If m_lngVisibleMessage >= lblMessage.Count Then
        btnNext.Enabled = False
    End If
End Sub

Private Sub swMessage(lngNum As Long)
    ' hide current message
    If m_lngVisibleMessage > -1 Then
        lblMessage(m_lngVisibleMessage).Visible = False
    End If
    
    ' show a message
    m_lngVisibleMessage = lngNum
    With lblMessage(lngNum)
        Caption = App.Title & " - " & .Tag
        .Visible = True
        .ZOrder 0
    End With
End Sub

Public Sub AddMessage(strMessage As String, strTitle As String)
    Dim X As Long

    X = lblMessage.Count
    If (X = 1) And (lblMessage(1).Tag = "") Then
        lblMessage(1).Caption = strMessage
        lblMessage(1).Tag = strTitle & " "
        swMessage 1
    Else
        X = X + 1
        Load lblMessage(X)
        lblMessage(X).Caption = strMessage
        lblMessage(X).Tag = strTitle & " "
        lblMessage(X).Visible = False
        btnOk.Left = 187
        btnNext.Visible = True
        btnNext.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If m_lngHIconLarge > 0 Then DestroyIcon m_lngHIconLarge
    If m_lngHIconSmall > 0 Then DestroyIcon m_lngHIconSmall
End Sub
