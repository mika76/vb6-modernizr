VERSION 5.00
Begin VB.UserControl ucList
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "ucList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  Custom-drawn list with shell file icons. Stock ListBoxes cannot be
'  owner-drawn after creation, so lists that want icons use this
'  control instead. API mirrors the ListBox members the forms already
'  use: Clear / AddItem / ListIndex / ListCount / Selected, plus
'  ItemKey for the parallel-array pattern and per-item icon or colored
'  dot. Long rows are clipped with an ellipsis instead of hscrolling.
'  Scrolling is a native WS_VSCROLL bar driven by modScroll (which
'  also delivers mouse-wheel input).
' =====================================================================

Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)

Private mText() As String
Private mKeys() As String
Private mIcon() As Long        ' cached HICON, 0 = none
Private mDot() As Long         ' dot color, -1 = none
Private mSel() As Boolean
Private mCount As Long

Private mListIndex As Long     ' focused row, -1 = none
Private mTop As Long           ' first visible row
Private mAnchor As Long        ' shift-range anchor
Private mMulti As Boolean
Private mHasGutter As Boolean  ' any icon/dot present -> indent text

Private mRowH As Long
Private mIconS As Long
Private mAttached As Boolean

' ---------------------------------------------------------------------
'  Public API
' ---------------------------------------------------------------------

Public Sub Clear()
    mCount = 0
    mListIndex = -1
    mAnchor = -1
    mTop = 0
    mHasGutter = False
    ReDim mText(0 To 32)
    ReDim mKeys(0 To 32)
    ReDim mIcon(0 To 32)
    ReDim mDot(0 To 32)
    ReDim mSel(0 To 32)
    UpdateScrollbar
    Repaint
End Sub

' iconFile: any path/name whose extension picks the shell icon.
' dotColor: small colored square instead of an icon (-1 = none).
Public Sub AddItem(ByVal sText As String, _
        Optional ByVal sKey As String = "", _
        Optional ByVal iconFile As String = "", _
        Optional ByVal dotColor As Long = -1)
    On Error Resume Next
    If mCount > UBound(mText) Then
        ReDim Preserve mText(0 To mCount * 2)
        ReDim Preserve mKeys(0 To mCount * 2)
        ReDim Preserve mIcon(0 To mCount * 2)
        ReDim Preserve mDot(0 To mCount * 2)
        ReDim Preserve mSel(0 To mCount * 2)
    End If
    mText(mCount) = sText
    mKeys(mCount) = sKey
    mIcon(mCount) = 0
    If Len(iconFile) > 0 Then mIcon(mCount) = IconForFile(iconFile)
    mDot(mCount) = dotColor
    mSel(mCount) = False
    If mIcon(mCount) <> 0 Or dotColor >= 0 Then mHasGutter = True
    mCount = mCount + 1
    UpdateScrollbar
    Repaint
End Sub

Public Property Get ListCount() As Long
    ListCount = mCount
End Property

Public Property Get ListIndex() As Long
    ListIndex = mListIndex
End Property

Public Property Let ListIndex(ByVal i As Long)
    If i < -1 Or i >= mCount Then Exit Property
    mListIndex = i
    mAnchor = i
    If Not mMulti Then SelectOnly i
    If i >= 0 Then EnsureVisible i
    Repaint
End Property

Public Property Get Selected(ByVal i As Long) As Boolean
    If i >= 0 And i < mCount Then Selected = mSel(i)
End Property

Public Property Let Selected(ByVal i As Long, ByVal v As Boolean)
    If i < 0 Or i >= mCount Then Exit Property
    mSel(i) = v
    Repaint
End Property

Public Function ItemKey(ByVal i As Long) As String
    If i >= 0 And i < mCount Then ItemKey = mKeys(i)
End Function

Public Function ItemText(ByVal i As Long) As String
    If i >= 0 And i < mCount Then ItemText = mText(i)
End Function

Public Property Get MultiSelect() As Boolean
    MultiSelect = mMulti
End Property

Public Property Let MultiSelect(ByVal v As Boolean)
    mMulti = v
End Property

Public Property Get TopIndex() As Long
    TopIndex = mTop
End Property

Public Property Let TopIndex(ByVal i As Long)
    If i < 0 Then i = 0
    If i > MaxTop() Then i = MaxTop()
    mTop = i
    UpdateScrollbar
    Repaint
End Property

' ---------------------------------------------------------------------
'  Lifecycle / layout
' ---------------------------------------------------------------------

Private Sub UserControl_Initialize()
    mRowH = ScaleForDpi(20)
    mIconS = ScaleForDpi(16)
    mListIndex = -1
    mAnchor = -1
    ReDim mText(0 To 32)
    ReDim mKeys(0 To 32)
    ReDim mIcon(0 To 32)
    ReDim mDot(0 To 32)
    ReDim mSel(0 To 32)
End Sub

Private Sub UserControl_Show()
    On Error Resume Next
    ' native scrollbar only at runtime, never in the form designer
    If Not mAttached And Ambient.UserMode Then
        mAttached = True
        Scroll_Attach UserControl.hwnd, Me
        UpdateScrollbar
    End If
End Sub

Private Sub UserControl_Terminate()
    On Error Resume Next
    If mAttached Then Scroll_Detach UserControl.hwnd
    mAttached = False
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    UpdateScrollbar
    Repaint
End Sub

' modScroll callback: scrollbar drag, arrows, page clicks, wheel
Public Sub ScrollTo(ByVal pos As Long)
    If pos < 0 Then pos = 0
    If pos > MaxTop() Then pos = MaxTop()
    If pos = mTop Then Exit Sub
    mTop = pos
    UpdateScrollbar
    Repaint
End Sub

Private Function VisRows() As Long
    If mRowH <= 0 Then mRowH = ScaleForDpi(20)
    VisRows = ScaleHeight \ mRowH
    If VisRows < 1 Then VisRows = 1
End Function

Private Function MaxTop() As Long
    MaxTop = mCount - VisRows()
    If MaxTop < 0 Then MaxTop = 0
End Function

Private Sub EnsureVisible(ByVal i As Long)
    If i < mTop Then mTop = i
    If i >= mTop + VisRows() Then mTop = i - VisRows() + 1
    If mTop < 0 Then mTop = 0
    UpdateScrollbar
End Sub

Private Sub UpdateScrollbar()
    On Error Resume Next
    If MaxTop() = 0 Then mTop = 0
    If mAttached Then _
        Scroll_Update UserControl.hwnd, mCount - 1, VisRows(), mTop
End Sub

' ---------------------------------------------------------------------
'  Selection / input
' ---------------------------------------------------------------------

Private Sub SelectOnly(ByVal i As Long)
    Dim j As Long
    For j = 0 To mCount - 1
        mSel(j) = (j = i)
    Next
End Sub

Private Sub SelectRange(ByVal a As Long, ByVal b As Long)
    Dim j As Long, lo As Long, hi As Long
    If a < 0 Then a = b
    lo = IIf(a < b, a, b)
    hi = IIf(a > b, a, b)
    For j = 0 To mCount - 1
        mSel(j) = (j >= lo And j <= hi)
    Next
End Sub

Private Function RowAt(ByVal y As Single) As Long
    Dim i As Long
    RowAt = -1
    If y < 0 Then Exit Function
    i = mTop + CLng(y) \ mRowH
    If i >= 0 And i < mCount And i < mTop + VisRows() Then RowAt = i
End Function

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, _
        x As Single, y As Single)
    On Error Resume Next
    Dim i As Long
    i = RowAt(y)
    If i < 0 Then Exit Sub

    If mMulti And (Shift And vbCtrlMask) <> 0 Then
        mSel(i) = Not mSel(i)
        mListIndex = i
        mAnchor = i
    ElseIf mMulti And (Shift And vbShiftMask) <> 0 Then
        SelectRange mAnchor, i
        mListIndex = i
    Else
        SelectOnly i
        mListIndex = i
        mAnchor = i
    End If
    Repaint
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim i As Long, moved As Boolean
    i = mListIndex

    Select Case KeyCode
        Case vbKeyUp:       i = i - 1: moved = True
        Case vbKeyDown:     i = i + 1: moved = True
        Case vbKeyPageUp:   i = i - (VisRows() - 1): moved = True
        Case vbKeyPageDown: i = i + (VisRows() - 1): moved = True
        Case vbKeyHome:     i = 0: moved = True
        Case vbKeyEnd:      i = mCount - 1: moved = True
    End Select

    If moved And mCount > 0 Then
        KeyCode = 0
        If i < 0 Then i = 0
        If i >= mCount Then i = mCount - 1
        mListIndex = i
        If mMulti And (Shift And vbShiftMask) <> 0 Then
            SelectRange mAnchor, i
        Else
            SelectOnly i
            mAnchor = i
        End If
        EnsureVisible i
        Repaint
        RaiseEvent Click
    Else
        RaiseEvent KeyDown(KeyCode, Shift)
    End If
End Sub

' ---------------------------------------------------------------------
'  Drawing
' ---------------------------------------------------------------------

Private Sub Repaint()
    On Error Resume Next
    Dim w As Long, i As Long, y As Long, lastRow As Long

    Cls
    w = ScaleWidth

    lastRow = mTop + VisRows() - 1
    If lastRow > mCount - 1 Then lastRow = mCount - 1

    For i = mTop To lastRow
        y = (i - mTop) * mRowH
        DrawRow i, y, w
    Next

    Refresh
End Sub

Private Sub DrawRow(ByVal i As Long, ByVal y As Long, ByVal w As Long)
    On Error Resume Next
    Dim xText As Long, cap As String, full As String

    If mSel(i) Then
        Line (0, y)-(w - 1, y + mRowH - 1), vbHighlight, BF
        ForeColor = vbHighlightText
    Else
        ForeColor = vbWindowText
    End If

    xText = ScaleForDpi(4)
    If mHasGutter Then
        If mIcon(i) <> 0 Then
            DrawIcon16 hdc, ScaleForDpi(4), y + (mRowH - mIconS) \ 2, mIcon(i)
        ElseIf mDot(i) >= 0 Then
            Line (ScaleForDpi(9), y + mRowH \ 2 - 3)- _
                 (ScaleForDpi(9) + 6, y + mRowH \ 2 + 3), mDot(i), BF
        End If
        xText = ScaleForDpi(4) + mIconS + ScaleForDpi(5)
    End If

    ' focused-but-unselected row gets a thin marker in the gutter
    If i = mListIndex And Not mSel(i) Then
        Line (0, y)-(1, y + mRowH - 1), vbHighlight, BF
    End If

    full = mText(i)
    cap = full
    Do While TextWidth(cap) > (w - xText - ScaleForDpi(6)) And Len(cap) > 1
        cap = Left$(cap, Len(cap) - 1)
    Loop
    If cap <> full Then cap = Left$(cap, Len(cap) - 1) & Chr$(133)

    CurrentX = xText
    CurrentY = y + (mRowH - TextHeight("X")) \ 2
    Print cap
End Sub
