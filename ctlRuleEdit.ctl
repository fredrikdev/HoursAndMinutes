VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.UserControl ctlRuleEdit 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2535
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   169
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox picSquare 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4305
      ScaleHeight     =   180
      ScaleWidth      =   300
      TabIndex        =   3
      Top             =   2160
      Width           =   300
   End
   Begin MSComCtl2.FlatScrollBar ctlScrollV 
      Height          =   1500
      Left            =   4200
      TabIndex        =   2
      Top             =   555
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2646
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1179648
   End
   Begin MSComCtl2.FlatScrollBar ctlScrollH 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Arrows          =   65536
      Orientation     =   1179649
   End
   Begin VB.PictureBox picContent 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   0
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   406
      TabIndex        =   0
      Top             =   0
      Width           =   6090
      Begin VB.Image imgHand 
         Height          =   480
         Left            =   2175
         Picture         =   "ctlRuleEdit.ctx":0000
         Top             =   1335
         Width           =   480
         Visible         =   0   'False
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   4
         Top             =   15
         Width           =   45
      End
   End
End
Attribute VB_Name = "ctlRuleEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' variables
Private m_lngSelectedHotspot As Long

' constants
Private Const m_intRowSpacing = 3

' events
Public Event LinkClick(ByVal strLinkID As String, ByRef strValue As String, ByRef blnSetValue As Boolean)

Private Sub ctlScrollH_Change()
    picContent.Left = -ctlScrollH.Value
End Sub

Private Sub ctlScrollV_Change()
    picContent.Top = -ctlScrollV.Value
End Sub

Private Sub UserControl_Resize()
    picContent.Move 0, 0
    showScroll
End Sub

' handles a click on a text
Private Sub lblText_Click(Index As Integer)
    Dim blnSetValue As Boolean, strValue As String
    
    lblText_MouseMove Index, 0, 0, 0, 0
    
    blnSetValue = False
    strValue = lblText(Index).Caption
    RaiseEvent LinkClick(lblText(Index).Tag, strValue, blnSetValue)
    If blnSetValue = True Then
        lblText(Index).Caption = FormatCaption(strValue)
        lblText(Index).AutoSize = True
        lblText(Index).Width = lblText(Index).Width + 2
        lblText(Index).Height = lblText(Index).Height + 1
        lblText(Index).Left = lblText(Index - 1).Left + lblText(Index - 1).Width
        lblText(Index + 1).Left = lblText(Index).Left + lblText(Index).Width
        showScroll
    End If
End Sub

' hilights a hotspot
Private Sub lblText_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    With lblText(Index)
        If .FontUnderline = False Then Exit Sub
        If Index = m_lngSelectedHotspot Then Exit Sub
        
        If m_lngSelectedHotspot >= 0 Then
            lblText(m_lngSelectedHotspot).BackColor = vbWindowBackground
            lblText(m_lngSelectedHotspot).ForeColor = vbButtonText
        End If
        .BackColor = vbHighlight
        .ForeColor = vbWindowBackground
        m_lngSelectedHotspot = Index
    End With
End Sub


' sets the text (and hotspots) for this control. clicks are reported through
' the LinkClick event
Public Sub setText(strText As String)
    ' |link id*value|
    Dim arrLines() As String, x As Long, y As Long
    Dim lngTop As Long, lngLeft As Long
    
    ' clear selected hotspot
    m_lngSelectedHotspot = -1
    
    ' unload previously set text
    For x = lblText.Count - 1 To 1 Step -1
        Unload lblText(x)
    Next
    lblText(0).Visible = False
    
    ' load new text
    arrLines = Split(strText, vbTab)
    y = 0
    For x = LBound(arrLines) To UBound(arrLines)
        lngLeft = lblText(0).Left
        If y <> 0 Then
            lngTop = lblText(y - 1).Top + lblText(y - 1).Height + m_intRowSpacing
            Load lblText(y)
        Else
            lngTop = lblText(0).Top
        End If
            
        If InStr(1, arrLines(x), "|") > 0 Then
            ' line with expression
            lblText(y).Move lngLeft, lngTop
            lblText(y).Caption = Left(arrLines(x), InStr(1, arrLines(x), "|") - 1)
            lblText(y).Visible = True
            y = y + 1
            
            Load lblText(y)
            lblText(y).Move lblText(y - 1).Left + lblText(y - 1).Width, lngTop
            lblText(y).Caption = FormatCaption(Mid(arrLines(x), InStr(1, arrLines(x), "*") + 1, InStrRev(arrLines(x), "|") - InStr(1, arrLines(x), "*") - 1))
            lblText(y).Tag = Mid(arrLines(x), InStr(1, arrLines(x), "|") + 1, InStr(1, arrLines(x), "*") - InStr(1, arrLines(x), "|") - 1)
            lblText(y).Alignment = vbCenter
            lblText(y).FontUnderline = True
            lblText(y).MousePointer = vbCustom
            lblText(y).MouseIcon = imgHand.Picture
            lblText(y).Height = lblText(y).Height + 1
            lblText(y).Width = lblText(y).Width + 2
            lblText(y).ZOrder 0
            lblText(y).Visible = True
            y = y + 1
            
            Load lblText(y)
            lblText(y).Move lblText(y - 1).Left + lblText(y - 1).Width, lngTop
            lblText(y).Caption = Mid(arrLines(x), InStrRev(arrLines(x), "|") + 1)
            lblText(y).Visible = True
        Else
            ' simple line
            lblText(y).Move lngLeft, lngTop
            lblText(y).Caption = arrLines(x)
            lblText(y).Visible = True
        End If
        
        y = y + 1
    Next
    showScroll
End Sub


' sizes scrollbars
Private Sub showScroll()
    Dim blnShowH As Boolean, blnShowV As Boolean, x As Long
    Dim lngMaxW As Long, lngMaxH As Long
    
    lngMaxW = -1
    lngMaxH = -1
    For x = 0 To lblText.Count - 1
        If lblText(x).Width + lblText(x).Left > lngMaxW Then lngMaxW = lblText(x).Width + lblText(x).Left
        If lblText(x).Height + lblText(x).Top > lngMaxH Then lngMaxH = lblText(x).Height + lblText(x).Top
    Next
    picContent.Width = lngMaxW
    picContent.Height = lngMaxH
    
    
    If picContent.Width > UserControl.ScaleWidth Then blnShowH = True
    If picContent.Height > UserControl.ScaleHeight Then blnShowV = True
    
    If blnShowH = True And blnShowV = True Then
        picSquare.Width = ctlScrollV.Width
        picSquare.Height = ctlScrollH.Height
        picSquare.Move UserControl.ScaleWidth - picSquare.Width, UserControl.ScaleHeight - picSquare.Height
    End If
    
    If blnShowH Then
        If blnShowV Then
            ctlScrollH.Move 0, UserControl.ScaleHeight - ctlScrollH.Height, UserControl.ScaleWidth - picSquare.Width, ctlScrollH.Height
        Else
            ctlScrollH.Move 0, UserControl.ScaleHeight - ctlScrollH.Height, UserControl.ScaleWidth, ctlScrollH.Height
        End If
        ctlScrollH.SmallChange = picContent.TextWidth("A")
        ctlScrollH.LargeChange = ctlScrollH.SmallChange * 2
        ctlScrollH.max = picContent.Width - UserControl.ScaleWidth + ctlScrollH.Height
        ctlScrollH.Visible = True
    Else
        ctlScrollH.Visible = False
    End If
    
    If blnShowV Then
        If blnShowH Then
            ctlScrollV.Move UserControl.ScaleWidth - ctlScrollV.Width, 0, ctlScrollV.Width, UserControl.ScaleHeight - picSquare.Height
        Else
            ctlScrollV.Move UserControl.ScaleWidth - ctlScrollV.Width, 0, ctlScrollV.Width, UserControl.ScaleHeight
        End If
        ctlScrollV.SmallChange = lblText(0).Height + m_intRowSpacing
        ctlScrollV.LargeChange = ctlScrollV.SmallChange * 2
        ctlScrollV.max = picContent.Height - UserControl.ScaleHeight + ctlScrollV.Width
        ctlScrollV.Visible = True
    Else
        ctlScrollV.Visible = False
    End If
    
    If blnShowH = True And blnShowV = True Then
        picSquare.Visible = True
    Else
        picSquare.Visible = False
    End If
End Sub

Private Function FormatCaption(strText As String) As String
    If InStr(1, strText, vbCrLf) Then
        FormatCaption = Left(strText, InStr(1, strText, vbCrLf) - 1) & "..."
        Exit Function
    End If
    
    If Len(strText) > 40 Then
        FormatCaption = Left(strText, 40) & "..."
        Exit Function
    End If
    
    FormatCaption = strText
End Function
