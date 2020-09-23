VERSION 5.00
Begin VB.UserControl CIMenu 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000010&
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   FillStyle       =   0  'Solid
   PropertyPages   =   "IMenu.ctx":0000
   ScaleHeight     =   326
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   280
   Begin VB.PictureBox P 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   3345
      Left            =   1140
      ScaleHeight     =   223
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   106
      TabIndex        =   0
      Top             =   900
      Width           =   1590
   End
End
Attribute VB_Name = "CIMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private Enum eEdge
    eLeft = 1
    eTop = 2
    eRight = 4
    eBottom = 8
    eTopleft = 3
    eBottomRight = 12
    eRect = 15
End Enum
Private Enum eButtonState
    eOver
    eDown
    eNormal
End Enum

Dim prevOverItem As Long
Dim prevOver As Long
Dim CMenuItems As New Collection

'Default Property Values:
Const m_def_ItemHeight = 60
Const m_def_Selected = 0
Const m_def_ButtonHeight = 20
'Property Variables:
Dim m_ImageList As Object
Dim m_ItemHeight As Long
Dim m_Selected As Long
Dim m_ButtonHeight As Long
Dim m_ItmDown As Long

Public Event ItemClick(Id As Long, Button As Integer, Shift As Integer)


Public Function SubItem() As CMenuItem
On Error Resume Next
Set SubItem = CMenuItems(Selected)
End Function

Public Function Count() As Long: Count = CMenuItems.Count: End Function
Public Sub Add(Caption As String, Optional Key As Variant)
    Dim Itm As New CMenuItem
    Itm.Caption = Caption
    If IsMissing(Key) = False Then CMenuItems.Add Itm, Key Else CMenuItems.Add Itm
    Redraw
End Sub
Public Sub Remove(Key As Variant)
    On Error Resume Next
    CMenuItems.Remove Key
    Redraw
    If m_Selected > Count Then Selected = Count
End Sub
Public Function Item(Key As Variant) As CMenuItem
    If Key > 0 And Key <= Count Then Set Item = CMenuItems(Key)
End Function
Public Sub Clear()
    Set CMenuItems = Nothing: m_Selected = 0
End Sub
Public Function Caption(Index As Variant) As String
    If Index > 0 And Index <= Count Then Caption = Item(Index).Caption
End Function
Public Function ItemCount(Index As Variant) As Long
    On Error Resume Next
    If Index > 0 And Index <= Count Then ItemCount = Item(Index).Count
End Function
Public Property Get TopItem() As Long
Attribute TopItem.VB_ProcData.VB_Invoke_Property = "Page1"
    If Selected > 0 And Selected <= Count Then TopItem = Item(Selected).TopItem
End Property
Public Property Let TopItem(Value As Long)
    Item(Selected).TopItem = Value: DrawItems
End Property
Private Sub DrawLine(Dc As Long, Left As Long, Top As Long, Right As Long, Bottom As Long)
   MoveToEx Dc, Left, Top, ByVal 0
   LineTo Dc, Right, Bottom
End Sub
Private Function PointInRect(x As Variant, y As Variant, R As RECT) As Boolean
    If x >= R.Left And x <= R.Right And y >= R.Top And y <= R.Bottom Then PointInRect = True
End Function
Private Sub PrintAt(Dc As Long, x As Long, y As Long, ByVal Str As String)
    TextOut Dc, x, y, Str, Len(Str)
End Sub
Private Sub ClearRect(Dc As Long, lpRect As RECT)
    With lpRect: Rectangle Dc, .Left, .Top, .Right, .Bottom + 1: End With
End Sub
Private Function ButtonRect(Index As Long) As RECT
On Error Resume Next
Dim ItmCount As Long, Ips As Long

ItmCount = Item(Index).Count
Ips = P.ScaleHeight \ m_ItemHeight

With ButtonRect
.Left = 1
.Right = ScaleWidth - 2
If Index = m_Selected And Ips < ItmCount Then
    .Right = .Right - m_ButtonHeight * 0.7
    .Left = .Left
Else
    Item(Index).TopItem = 0
End If
If Index <= m_Selected Then
    .Top = 1 + (Index - 1) * (m_ButtonHeight + 1)
    .Bottom = .Top + m_ButtonHeight
Else
    .Top = ScaleHeight - (Count - Index + 1) * (m_ButtonHeight + 1) - 1
    .Bottom = .Top + m_ButtonHeight
End If
End With
End Function
Private Function ItemRect(Index As Long) As RECT
    With ItemRect
        .Left = 1
        .Right = P.ScaleWidth - 1
        .Top = (Index - 1 - TopItem) * (m_ItemHeight + 1)
        .Bottom = .Top + m_ItemHeight
    End With
End Function
Private Function ItemsRect() As RECT
With ItemsRect
    .Left = 2
    .Right = ScaleWidth - 2
If Count <> 0 Then
    .Top = ButtonRect(m_Selected).Bottom + 1
    .Bottom = ButtonRect(m_Selected + 1).Top - 1
Else
    .Top = 1
    .Bottom = ScaleHeight - 2
End If
End With
End Function
Private Function ScrollRect(Index As Long) As RECT
    Dim R As RECT
    R = ButtonRect(m_Selected)
           ScrollRect.Left = R.Right + 1
           ScrollRect.Right = ScaleWidth - 2
    Select Case Index
        Case Is = 1
           ScrollRect.Top = R.Top
           ScrollRect.Bottom = R.Top + m_ButtonHeight / 2 - 1
        Case Is = 2
           ScrollRect.Top = R.Top + m_ButtonHeight / 2
           ScrollRect.Bottom = R.Bottom
    End Select
End Function

Public Sub Redraw()

UserControl.Cls
Dim I As Long

If m_Selected = 0 And Count > 0 Then
    m_Selected = 1
End If
DrawItems

If CMenuItems Is Nothing Then Exit Sub

For I = 1 To Count
    DrawButton I, eNormal
Next

DrawScroll 1, eNormal
DrawScroll 2, eNormal

Dim iR As RECT

SetRect iR, 0, 0, ScaleWidth - 1, ScaleHeight - 1
DrawEdge hdc, iR.Left, iR.Top, iR.Right, iR.Bottom, eBottomRight, vb3DHighlight

iR = ItemsRect
DrawEdge hdc, iR.Left - 1, iR.Top, iR.Right, iR.Bottom, eTopleft, vb3DDKShadow

If AutoRedraw Then Refresh
End Sub


Private Sub DrawButton(Index As Long, State As eButtonState)
    DrawButtonEx ButtonRect(Index), Caption(Index), State
End Sub


Private Sub DrawButtonEx(mR As RECT, Caption As String, State As eButtonState)

On Error Resume Next
Dim cX As Long
Dim cY As Long
With mR
If State = eNormal Then
    FontBold = False
Else
    FontBold = True
End If
cX = (.Right + .Left - TextWidth(Caption)) / 2
cY = (.Bottom + .Top - TextHeight("I")) / 2
FillColor = vbButtonFace
ForeColor = vbButtonFace
ClearRect hdc, mR

Select Case State
 Case eNormal, eUp
  DrawEdge hdc, .Left, .Top, .Right, .Bottom, eBottomRight, vbButtonShadow
  DrawEdge hdc, .Left, .Top, .Right, .Bottom, eTopleft, vb3DHighlight
  UserControl.ForeColor = vbButtonText
  PrintAt hdc, cX, cY, Caption
 Case Is = eDown
  DrawEdge hdc, .Left, .Top, .Right, .Bottom, eBottomRight, vb3DHighlight
  DrawEdge hdc, .Left, .Top, .Right, .Bottom, eTopleft, vbButtonShadow
  UserControl.ForeColor = vbButtonText
  PrintAt hdc, cX + 1, cY + 1, Caption
End Select
End With
End Sub
Private Sub DrawItems()
    On Error Resume Next
    If m_Selected < 1 And m_Selected > Count Then Exit Sub
    
    Dim I As Long
    Dim R As RECT
    
    R = ItemsRect
    P.Left = R.Left
    P.Top = R.Top + 1
    P.Width = R.Right - R.Left
    P.Height = R.Bottom - R.Top - 1
    
    P.Cls
    For I = 1 To ItemCount(m_Selected)
        DrawItem I, eNormal
    Next
End Sub

Private Sub DrawScroll(Index As Long, State As eButtonState)
    Dim R As RECT
    Dim rWidth As Long
    Dim x As Long
    Dim y As Long
    Dim pFont As String
    
    R = ScrollRect(Index)
    rWidth = Abs(R.Right - R.Left)
    
    If rWidth < 3 Then Exit Sub
    DrawButtonEx R, "", State
    
    pFont = FontName
    FontName = "Webdings"
    
    x = (R.Left + R.Right - TextWidth("5")) / 2 + 1
    y = (R.Top + R.Bottom - TextHeight("5")) / 2 - 1
    
    If Index = 1 Then
        PrintAt hdc, x, y, "5"
    Else
        PrintAt hdc, x, y, "6"
    End If
    FontName = pFont
End Sub

Private Sub DrawItem(Index As Long, State As eButtonState)
On Error Resume Next

Dim R As RECT
R = ItemRect(Index)

Dim Itm As CItem
Set Itm = Item(m_Selected).Item(Index)

Dim cX As Long
Dim cY As Long

P.FillColor = vbButtonShadow
R.Right = R.Right + 1
ClearRect P.hdc, R
R.Right = R.Right - 1

FontBold = False
cX = (R.Left + R.Right - TextWidth(Itm.Caption)) / 2
cY = R.Bottom - TextHeight("I") - 2
Dim Color As Long
Color = P.ForeColor
P.ForeColor = vbWhite
PrintAt P.hdc, cX, cY, Itm.Caption
P.ForeColor = Color

Select Case State
    Case Is = eDown
        DrawEdge P.hdc, R.Left, R.Top, R.Right, R.Bottom, eTopleft, vb3DDKShadow
        DrawEdge P.hdc, R.Left, R.Top, R.Right, R.Bottom, eBottomRight, vbButtonFace
    Case Is = eOver
        DrawEdge P.hdc, R.Left, R.Top, R.Right, R.Bottom, eTopleft, vbButtonFace
        DrawEdge P.hdc, R.Left, R.Top, R.Right, R.Bottom, eBottomRight, vb3DDKShadow
    Case Is = eNormal
        DrawEdge P.hdc, R.Left, R.Top, R.Right, R.Bottom, eRect, vb3DShadow
End Select

cX = (R.Left + R.Right - m_ImageList.ImageWidth) / 2
cY = (R.Top + R.Bottom - m_ImageList.ImageHeight - TextHeight("I") - 3) / 2

If Itm.Icon <> 0 Then
P.PaintPicture m_ImageList.ListImages(Itm.Icon).ExtractIcon, cX, cY
End If

End Sub

Private Sub DrawEdge(Dc As Long, Left As Long, Top As Long, Right As Long, Bottom As Long, F As eEdge, ByVal Color As Long)
    Dim hPen As Long
    Dim pPen As Long
    TranslateColor Color, 0, Color
    hPen = CreatePen(0, 1, Color)
    pPen = SelectObject(Dc, hPen)
    
    If F And eLeft Then DrawLine Dc, Left, Top, Left, Bottom
    If F And eTop Then DrawLine Dc, Left, Top, Right, Top
    If F And eRight Then DrawLine Dc, Right, Top, Right, Bottom + 1
    If F And eBottom Then DrawLine Dc, Left, Bottom, Right, Bottom
    
    DeleteObject SelectObject(Dc, pPen)
End Sub

Public Property Get ButtonHeight() As Long
Attribute ButtonHeight.VB_ProcData.VB_Invoke_Property = "Page1"
    ButtonHeight = m_ButtonHeight
End Property
Public Property Let ButtonHeight(ByVal New_ButtonHeight As Long)
    m_ButtonHeight = New_ButtonHeight
    PropertyChanged "ButtonHeight"
End Property





Private Sub P_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim I As Long
    
      For I = 1 To ItemCount(m_Selected)
        If PointInRect(x, y, ItemRect(I)) Then
           If Button = 1 Then DrawItem I, eDown
           
           m_ItmDown = I
           Exit Sub
         End If
      Next
End Sub

Private Sub P_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    SetCapture P.hwnd
    
    Dim CPos As POINTAPI
    Dim R As RECT
    Dim I As Long
    Dim InRect As Byte
    
    For I = 1 To ItemCount(m_Selected)
      R = ItemRect(I)
      If R.Bottom > P.ScaleHeight Then
         R.Bottom = P.ScaleHeight - 1
      End If
        If PointInRect(x, y, R) Then
         InRect = 1
         If I <> prevOverItem Then
           DrawItem prevOverItem, eNormal
           If Button <> 1 Then
             DrawItem I, eOver
           Else
             DrawItem I, eDown
           End If
           prevOverItem = I
         End If
       End If
    Next
    If InRect <> 1 Then
      DrawItem prevOverItem, eNormal
      prevOverItem = 0
    End If
    
    GetCursorPos CPos
    
    Dim hWndOver As Long
    hWndOver = WindowFromPoint(CPos.x, CPos.y)
    If hWndOver <> P.hwnd Then
        ReleaseCapture
    End If
End Sub

Private Sub P_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   DrawItem prevOverItem, eNormal
   If PointInRect(x, y, ItemRect(m_ItmDown)) Then
      RaiseEvent ItemClick(m_ItmDown, Button, Shift)
      If Button = 2 Then DrawItem m_ItmDown, eNormal
      prevOverItem = 0
   End If
End Sub

Private Sub UserControl_DblClick()
'To Do:*
End Sub



Private Sub UserControl_InitProperties()
    m_ButtonHeight = m_def_ButtonHeight
    m_Selected = m_def_Selected
    m_ItemHeight = m_def_ItemHeight
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim I As Long
    If Button = 1 Then
      For I = 1 To Count
        If PointInRect(x, y, ButtonRect(I)) Then
           DrawButton I, eDown
           Exit Sub
        End If
      Next
      For I = 1 To 2
        If PointInRect(x, y, ScrollRect(I)) Then
          DrawScroll I, eDown
        End If
      Next
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim R As RECT
    Dim mX As Long
    Dim mY As Long
    
    mX = CLng(x)
    mY = CLng(y)
    
    
    SetCapture hwnd
    R.Bottom = ScaleHeight
    R.Right = ScaleWidth
    
    Dim hWndOver As Long
    
    Dim CPos As POINTAPI
    GetCursorPos CPos
    
    hWndOver = WindowFromPoint(CPos.x, CPos.y)
    
    Dim I As Long
    Dim InRect As Byte
    
    For I = 1 To Count
        R = ButtonRect(I)
        If PointInRect(mX, mY, R) Then
          InRect = 1
          If I <> prevOver Then
            DrawButton prevOver, eNormal
            If Button <> 1 Then
              DrawButton I, eOver
            Else
              DrawButton I, eDown
            End If
            prevOver = I
           End If
        End If
    Next
    If InRect <> 1 Then
            DrawButton prevOver, eNormal
            prevOver = 0
    End If
    For I = 1 To 2
        If PointInRect(x, y, ScrollRect(I)) Then
            If Button = 1 Then
                DrawScroll I, eDown
            Else
                DrawScroll I, eNormal
            End If
        Else
            DrawScroll I, eNormal
        End If
    Next
    If hWndOver <> hwnd Then
        ReleaseCapture
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button <> 1 Then Exit Sub
   DrawButton prevOver, eOver

   If prevOver = m_Selected Then Exit Sub

   Dim R As RECT
   R = ButtonRect(prevOver)
   If PointInRect(x, y, R) Then
     Selected = prevOver
   End If
   
   DrawScroll 1, eNormal
   DrawScroll 2, eNormal
   
   Dim Ips As Long
   Ips = P.ScaleHeight \ m_ItemHeight

   If PointInRect(x, y, ScrollRect(2)) Then
     If TopItem + Ips < ItemCount(m_Selected) Then TopItem = TopItem + 1
   End If
   If PointInRect(x, y, ScrollRect(1)) Then
     If TopItem > 0 Then TopItem = TopItem - 1
   End If
   
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_ButtonHeight = PropBag.ReadProperty("ButtonHeight", m_def_ButtonHeight)
    m_Selected = PropBag.ReadProperty("Selected", m_def_Selected)
    m_ItemHeight = PropBag.ReadProperty("ItemHeight", m_def_ItemHeight)
    Set m_ImageList = PropBag.ReadProperty("ImageList", Nothing)
End Sub

Private Sub UserControl_Resize()
    Redraw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ButtonHeight", m_ButtonHeight, m_def_ButtonHeight)
    Call PropBag.WriteProperty("Selected", m_Selected, m_def_Selected)
    Call PropBag.WriteProperty("ItemHeight", m_ItemHeight, m_def_ItemHeight)
    Call PropBag.WriteProperty("ImageList", m_ImageList, Nothing)
End Sub

Public Property Get Selected() As Long
Attribute Selected.VB_ProcData.VB_Invoke_Property = "Page1"
    Selected = m_Selected
End Property

Public Property Let Selected(ByVal New_Selected As Long)
    If New_Selected <> m_Selected Then
        m_Selected = New_Selected
        Redraw
    End If
    PropertyChanged "Selected"
End Property

Public Property Get ItemHeight() As Long
Attribute ItemHeight.VB_ProcData.VB_Invoke_Property = "Page1"
    ItemHeight = m_ItemHeight
End Property

Public Property Let ItemHeight(ByVal New_ItemHeight As Long)
    m_ItemHeight = New_ItemHeight
    PropertyChanged "ItemHeight"
End Property

Public Property Get ImageList() As Object
    Set ImageList = m_ImageList
End Property

Public Property Set ImageList(ByVal New_ImageList As Object)
    Set m_ImageList = New_ImageList
    PropertyChanged "ImageList"
End Property


Public Property Get SelectedItem() As Variant
    SelectedItem = m_ItmDown
End Property

Public Sub SetItemCaption(Index As Long, NewCaption As String)
        
    SubItem.Item(Index).Caption = NewCaption
    DrawItem Index, eNormal
End Sub

