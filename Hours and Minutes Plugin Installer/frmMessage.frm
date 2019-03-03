VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hours and Minutes Plugin Installer"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
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
   Icon            =   "frmMessage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   122
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOk 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   345
      Left            =   4170
      TabIndex        =   1
      Top             =   1365
      Width           =   1125
   End
   Begin VB.Label lblCaption 
      Height          =   1065
      Left            =   165
      TabIndex        =   0
      Top             =   150
      Width           =   5130
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lngExitCount As Long

Private Sub btnOk_Click()
    End
End Sub
