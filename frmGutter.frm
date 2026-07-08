VERSION 5.00
Begin VB.Form frmGutter
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000F&
   BorderStyle     =   0  'None
   Caption         =   ""
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   600
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   40
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmGutter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  Line-number gutter: a strip reserved at the left of the MDI client
'  (same mechanism the tab bar uses for the top). Draws the CodeModule
'  line numbers of the ACTIVE code pane, aligned with its rows. Only a
'  maximized MDI child lines up with the strip - floating windows fall
'  back to the small in-margin numbers (see modHighlight).
' =====================================================================

Private mSig As String      ' last drawn state, skips no-op repaints

Public Sub Attach()
    On Error Resume Next
    If GetParent(Me.hwnd) = MainHwnd() Then Exit Sub
    Dim style As Long
    style = GetWindowLongA(Me.hwnd, GWL_STYLE)
    style = (style Or WS_CHILD) And (Not WS_POPUP)
    SetWindowLongA Me.hwnd, GWL_STYLE, style
    SetParent Me.hwnd, MainHwnd()
    mSig = ""
    Poll
End Sub

Public Sub Reposition()
    On Error Resume Next
    If Not gLineNumsEnabled Then Exit Sub
    Dim hMDI As Long, rc As RECT, pt As POINTAPI
    hMDI = MDIClientHwnd()
    If hMDI = 0 Then Exit Sub
    GetWindowRect hMDI, rc
    pt.x = rc.Left: pt.y = rc.Top
    ScreenToClient MainHwnd(), pt
    MoveWindow Me.hwnd, pt.x - gReserveLeftPx, pt.y, _
               gReserveLeftPx, rc.Bottom - rc.Top, 1
    mSig = ""
    Poll
End Sub

' Repaint only when the visible state changed. Called from the pane
' subclass (paint/scroll) and the tab-bar timer (pane switches).
Public Sub Poll()
    On Error Resume Next
    If Not Me.Visible Then Exit Sub

    Dim cp As VBIDE.CodePane, hPane As Long, sig As String
    Set cp = gVBE.ActiveCodePane
    If Not cp Is Nothing Then hPane = CodePaneHwnd(cp)

    If hPane <> 0 Then
        If (GetWindowLongA(hPane, GWL_STYLE) And WS_MAXIMIZE) = 0 Then _
            hPane = 0
    End If

    If hPane = 0 Then
        If mSig <> "-" Then
            mSig = "-"
            Me.Cls
            DrawEdge
            Me.Refresh
        End If
        Exit Sub
    End If

    sig = cp.CodeModule.Parent.Name & "|" & cp.topLine & "|" & _
          cp.CountOfVisibleLines & "|" & cp.CodeModule.CountOfLines & _
          "|" & EditorTopOffset(hPane) & "|" & PaneOffsetY(hPane)
    If sig = mSig Then Exit Sub
    mSig = sig

    DrawNumbers cp, hPane
End Sub

' ---------------------------------------------------------------------

Private Sub DrawNumbers(cp As VBIDE.CodePane, ByVal hPane As Long)
    On Error Resume Next
    Me.Cls

    Dim rc As RECT, yTop As Long, visLines As Long, topLine As Long
    Dim lineH As Long, cnt As Long
    GetClientRect hPane, rc
    yTop = EditorTopOffset(hPane)
    visLines = cp.CountOfVisibleLines
    topLine = cp.topLine
    cnt = cp.CodeModule.CountOfLines
    If visLines < 1 Then DrawEdge: Me.Refresh: Exit Sub
    lineH = (rc.Bottom - rc.Top - yTop) \ visLines
    If lineH < 4 Then DrawEdge: Me.Refresh: Exit Sub

    ' vertical offset of the pane's client area within the gutter;
    ' text rows start a hair below the computed offset (same empirical
    ' shift the match boxes use in modHighlight)
    Dim dy As Long
    dy = PaneOffsetY(hPane) + yTop + ScaleForDpi(3)

    ' in Procedure View only the current proc's lines are displayed
    Dim loLine As Long, hiLine As Long
    GetDisplayedRange cp, loLine, hiLine

    ' font sized to the editor's line height
    Dim dpi As Long, fpx As Long
    dpi = GetDeviceCaps(Me.hdc, LOGPIXELSY)
    If dpi <= 0 Then dpi = 96
    fpx = lineH - ScaleForDpi(5)
    If fpx < ScaleForDpi(8) Then fpx = ScaleForDpi(8)
    Me.FontName = "Segoe UI"
    Me.FontSize = fpx * 72 / dpi
    Me.ForeColor = vbGrayText

    Dim i As Long, ln As Long, s As String, y As Long, xRight As Long
    xRight = Me.ScaleWidth - ScaleForDpi(6)
    For i = 0 To visLines - 1
        ln = topLine + i
        If ln >= loLine And ln <= hiLine And ln <= cnt Then
            y = dy + i * lineH
            If y + lineH > Me.ScaleHeight Then Exit For
            If y >= 0 Then
                s = CStr(ln)
                Me.CurrentX = xRight - Me.TextWidth(s)
                Me.CurrentY = y + (lineH - Me.TextHeight(s)) \ 2
                Me.Print s
            End If
        End If
    Next

    DrawEdge
    Me.Refresh
End Sub

Private Sub DrawEdge()
    On Error Resume Next
    Me.Line (Me.ScaleWidth - 1, 0)-(Me.ScaleWidth - 1, Me.ScaleHeight), _
        vb3DShadow
End Sub

' Top of the pane's client area relative to the gutter's client area.
Private Function PaneOffsetY(ByVal hPane As Long) As Long
    On Error Resume Next
    Dim p0 As POINTAPI, p1 As POINTAPI
    p0.x = 0: p0.y = 0
    ClientToScreen hPane, p0
    p1.x = 0: p1.y = 0
    ClientToScreen Me.hwnd, p1
    PaneOffsetY = p0.y - p1.y
End Function
