VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl ctlCommentEdit 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5160
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   1545
   ScaleWidth      =   5160
   Begin VB.PictureBox picContent 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2565
      Left            =   0
      ScaleHeight     =   2565
      ScaleWidth      =   3345
      TabIndex        =   1
      Top             =   0
      Width           =   3345
      Begin VB.Shape shpSelection 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H8000000D&
         FillStyle       =   0  'Solid
         Height          =   285
         Left            =   570
         Top             =   660
         Width           =   1215
         Visible         =   0   'False
      End
      Begin VB.Label lblText 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This is a jTest label"
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   1350
      End
   End
   Begin MSComCtl2.FlatScrollBar ctlScroll 
      Height          =   1500
      Left            =   4890
      TabIndex        =   0
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2646
      _Version        =   393216
      Appearance      =   0
      Orientation     =   1179648
   End
   Begin VB.Image imgBG 
      Height          =   270
      Left            =   3570
      Picture         =   "ctlCommentEdit.ctx":0000
      Top             =   1080
      Width           =   270
      Visible         =   0   'False
   End
End
Attribute VB_Name = "ctlCommentEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lngLastItem As Long
Private m_blnResizing As Boolean
Private m_lngVisibleItems As Long

Public Event DblClick(ByRef strText As String, ByRef blnUpdate As Boolean, ByRef blnAddItem As Boolean)
Public Event Delete(ByRef strText As String, ByRef blnDelete As Boolean)

Private Sub hilightItem(lngItem As Long)
    Dim y As Long
    
    ' check if same
    If lngItem = m_lngLastItem Then Exit Sub
    
    ' hilight item
    If lblText.Count > lngItem Then
        With lblText(lngItem)
            .ForeColor = vbWhite
        End With
    Else
        Exit Sub
    End If
    
    ' remove old hilight
    If m_lngLastItem > -1 Then
        lblText(m_lngLastItem).ForeColor = vbBlack
    End If
        
    ' move selection
    y = lngItem * (19 * Screen.TwipsPerPixelY)
    shpSelection.Move 0, y, picContent.ScaleWidth + Screen.TwipsPerPixelX
    shpSelection.ZOrder 1
    shpSelection.Visible = True
    
    ' save settings
    m_lngLastItem = lngItem
End Sub

Private Sub ctlScroll_Change()
    picContent.Top = -(ctlScroll.Value * 19 * Screen.TwipsPerPixelY)
End Sub

Private Sub lblText_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    hilightItem CLng(Index)
End Sub

Private Sub picContent_DblClick()
    Dim strText As String, blnUpdate As Boolean, blnAddItem As Boolean
    
    strText = lblText(m_lngLastItem).Tag
    blnUpdate = False
    blnAddItem = False
    RaiseEvent DblClick(strText, blnUpdate, blnAddItem)
    
    If blnUpdate Then
        With lblText(m_lngLastItem)
            .Tag = strText
            .Caption = strText
            formatLabel m_lngLastItem
        End With
    ElseIf blnAddItem Then
        ' create new item
        
        Load lblText(lblText.Count)
        With lblText(lblText.Count - 1)
            .Tag = strText
            .Caption = strText
            .Move lblText(lblText.Count - 2).Left, lblText(lblText.Count - 2).Top + 19 * Screen.TwipsPerPixelY
            .Visible = True
            formatLabel lblText.Count - 1
        End With
        
        ' switch items (so that <new item> remains last)
        strText = lblText(lblText.Count - 1).Caption
        lblText(lblText.Count - 1).Caption = lblText(lblText.Count - 2).Caption
        lblText(lblText.Count - 2).Caption = strText
        
        strText = lblText(lblText.Count - 1).Tag
        lblText(lblText.Count - 1).Tag = lblText(lblText.Count - 2).Tag
        lblText(lblText.Count - 2).Tag = strText
        
        ' size content
        picContent.Height = ((Screen.TwipsPerPixelY * 19) * lblText.Count)
        
        ' fix
        lblText(lblText.Count - 1).ForeColor = vbBlack
        
        ' size scrollbar
        If lblText.Count > m_lngVisibleItems Then
            ctlScroll.Max = lblText.Count - m_lngVisibleItems
            ctlScroll.Enabled = True
        Else
            ctlScroll.Enabled = False
        End If
    End If
End Sub

Private Sub lblText_DblClick(Index As Integer)
    picContent_DblClick
End Sub

Private Sub picContent_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim z As Long, intItem As Long
    
    z = 0
    intItem = 0
    While z < y
        z = z + 19 * Screen.TwipsPerPixelY
        intItem = intItem + 1
    Wend
    intItem = intItem - 1
    z = z - 19 * Screen.TwipsPerPixelY
    If intItem < 0 Or z < 0 Then
        z = 0
        intItem = 0
    End If
    
    hilightItem intItem
End Sub

Private Sub UserControl_Initialize()
    m_lngLastItem = -1
    m_blnResizing = False
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngTemp As Long, lngVItems As Long
    Dim blnDelete As Boolean, x As Long
    
    lngVItems = m_lngVisibleItems - 1
    
    ' 40 = down, 38 = up
    If KeyCode = 40 Then
        ' down
        If m_lngLastItem + 1 > lblText.Count Then Exit Sub
        hilightItem m_lngLastItem + 1
    ElseIf KeyCode = 38 Then
        ' up
        If m_lngLastItem - 1 < 0 Then Exit Sub
        hilightItem m_lngLastItem - 1
    ElseIf KeyCode = 34 Then
        ' pgdn
        If m_lngLastItem + lngVItems >= lblText.Count Then
            lngTemp = lblText.Count - 1
        Else
            lngTemp = m_lngLastItem + lngVItems
        End If
        If lngTemp = m_lngLastItem Then Exit Sub
        
        hilightItem lngTemp
    ElseIf KeyCode = 33 Then
        ' pgup
        If m_lngLastItem - lngVItems < 0 Then
            lngTemp = 0
        Else
            lngTemp = m_lngLastItem - lngVItems
        End If
        If lngTemp = m_lngLastItem Then Exit Sub
        
        hilightItem lngTemp
    ElseIf KeyCode = 13 Then
        ' enter
        picContent_DblClick
    ElseIf KeyCode = 46 Then
        ' delete item
        blnDelete = False
        RaiseEvent Delete(lblText(m_lngLastItem).Tag, blnDelete)
        
        If blnDelete Then
            ' move items upwards
            For x = m_lngLastItem To lblText.Count - 2
                lblText(x).Caption = lblText(x + 1).Caption
                lblText(x).Tag = lblText(x + 1).Tag
            Next
            
            ' unload last item
            Unload lblText(lblText.Count - 1)
        
            ' size content
            picContent.Height = ((Screen.TwipsPerPixelY * 19) * lblText.Count)
            
            ' size scrollbar
            If lblText.Count > m_lngVisibleItems Then
                ctlScroll.Max = lblText.Count - m_lngVisibleItems
                ctlScroll.Enabled = True
            Else
                ctlScroll.Enabled = False
            End If
        End If
    End If
    
    ' check if selection is visible - else show it
    Do While (shpSelection.Top) + picContent.Top < 0
        If ctlScroll.Value - 1 >= 0 Then
            ctlScroll.Value = ctlScroll.Value - 1
        Else
            Exit Do
        End If
    Loop
    
    ' check if selection is visible - else show it
    Do While (shpSelection.Top) + picContent.Top > UserControl.ScaleHeight
        If ctlScroll.Value + 1 <= ctlScroll.Max Then
            ctlScroll.Value = ctlScroll.Value + 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Sub UserControl_Resize()
    Dim x As Long, y As Long
    
    If m_blnResizing = True Then Exit Sub
    m_blnResizing = True
    
    ' size main box
    x = 0
    While x < UserControl.ScaleHeight
        x = x + (19 * Screen.TwipsPerPixelY)
    Wend
    UserControl.Height = x + Screen.TwipsPerPixelY * 3
    
    ' size picture box
    picContent.Width = UserControl.ScaleWidth - ctlScroll.Width
    picContent.Left = 0
    
    ' fill with yellow background
    x = 0
    While x < picContent.ScaleWidth
        y = 0
        While y < picContent.ScaleHeight
            picContent.PaintPicture imgBG.Picture, x, y
            y = y + imgBG.Height
        Wend
        x = x + imgBG.Width
    Wend
    
    ' create separators
    y = imgBG.Height
    While y < picContent.ScaleHeight
        picContent.Line (0, y)-(picContent.ScaleWidth, y), vbButtonShadow
        y = y + imgBG.Height + Screen.TwipsPerPixelY
    Wend
    
    ' adjust scrollbar
    ctlScroll.Move UserControl.ScaleWidth - ctlScroll.Width, 0, ctlScroll.Width, UserControl.ScaleHeight
    
    m_blnResizing = False
End Sub

Public Function GetText() As String()
    Dim arrRet() As String
    ReDim arrRet(0 To lblText.Count - 1)
    Dim x As Long
    
    For x = 0 To lblText.Count - 1
        arrRet(x) = lblText(x).Tag
    Next
    
    GetText = arrRet
End Function

Public Sub SetText(arrText() As String)
    Dim x As Long, xx As Long, y As Long, z As Long
    
    For x = lblText.Count - 1 To 1 Step -1
        Unload lblText(x)
    Next
    
    ' load items
    y = lblText(0).Top
    xx = 0
    z = -1
    m_lngLastItem = -1
    For x = LBound(arrText) To UBound(arrText)
        If x <> 0 Then Load lblText(x)
        With lblText(x)
            .Caption = arrText(x)
            .Tag = arrText(x)
            .ForeColor = vbBlack
            formatLabel x
            
            .Move lblText(0).Left, y
            If (xx >= UserControl.ScaleHeight) And (z = -1) Then
                z = x
            End If
            .Visible = True
            xx = xx + 19 * Screen.TwipsPerPixelY
        End With
        y = y + 19 * Screen.TwipsPerPixelY
    Next
    picContent.Height = y - Screen.TwipsPerPixelY * 2
    
    ' determine maximum number of visible items
    xx = 0
    x = 0
    While xx < UserControl.ScaleHeight
        x = x + 1
        xx = xx + 19 * Screen.TwipsPerPixelY
    Wend
    m_lngVisibleItems = x
    
    hilightItem 0

    If z <> -1 Then
        ctlScroll.Max = UBound(arrText) - z + 1
        ctlScroll.Enabled = True
    Else
        ctlScroll.Enabled = False
    End If
    UserControl_Resize
End Sub

Private Sub formatLabel(lngIndex As Long)
    Dim blnCroped As Boolean
    
    blnCroped = False
    With lblText(lngIndex)
        If InStr(1, .Caption, vbCrLf) > 0 Then
            .Caption = Left(.Caption, InStr(1, .Caption, vbCrLf) - 1) & "..."
        End If
        While .Width + .Left + Screen.TwipsPerPixelX * 20 > UserControl.ScaleWidth - ctlScroll.Width
            .Caption = Left(.Caption, Len(.Caption) - 1)
            blnCroped = True
        Wend
        If (blnCroped = True) And (Right(.Caption, 3) <> "...") Then .Caption = .Caption & "..."
    End With
End Sub
