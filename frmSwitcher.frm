VERSION 5.00
Begin VB.Form frmSwitcher
   Appearance      =   0  'Flat
   BackColor       =   &H8000000F&
   BorderStyle     =   0  'None
   Caption         =   ""
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   5000
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstWins
      Height          =   2205
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5000
   End
End
Attribute VB_Name = "frmSwitcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  Ctrl+Tab MRU window switcher (Alt-Tab style). Shown WITHOUT taking
'  focus; while Ctrl is held, the message hook in modWheel drives it:
'  Ctrl+Tab / Ctrl+Shift+Tab step through the list, releasing Ctrl
'  activates the selection, Esc cancels.
' =====================================================================

Private mItemKeys() As String

Public Function BeginSwitch(ByVal forward As Boolean) As Boolean
    On Error Resume Next
    BuildList
    If lstWins.ListCount < 2 Then Exit Function

    SetWindowLongA Me.hwnd, GWL_HWNDPARENT, MainHwnd()
    SizeAndPosition
    lstWins.ListIndex = IIf(forward, 1, lstWins.ListCount - 1)
    gSwitcherActive = True
    ShowWindow Me.hwnd, SW_SHOWNOACTIVATE
    BeginSwitch = True
End Function

Public Sub StepSwitch(ByVal forward As Boolean)
    On Error Resume Next
    Dim i As Long
    i = lstWins.ListIndex + IIf(forward, 1, -1)
    If i < 0 Then i = lstWins.ListCount - 1
    If i >= lstWins.ListCount Then i = 0
    lstWins.ListIndex = i
End Sub

Public Sub CommitSwitch()
    On Error Resume Next
    Dim i As Long
    i = lstWins.ListIndex
    gSwitcherActive = False
    ShowWindow Me.hwnd, SW_HIDE
    If i >= 0 Then ActivateByKey mItemKeys(i)
End Sub

Public Sub CancelSwitch()
    On Error Resume Next
    gSwitcherActive = False
    ShowWindow Me.hwnd, SW_HIDE
End Sub

' clicking an entry with the mouse also commits
Private Sub lstWins_MouseUp(Button As Integer, Shift As Integer, _
        x As Single, y As Single)
    If gSwitcherActive Then CommitSwitch
End Sub

' ---------------------------------------------------------------------

Private Sub BuildList()
    On Error Resume Next
    Dim w As VBIDE.Window, k As String, cap As String, i As Long

    lstWins.Clear
    ReDim mItemKeys(0 To 32)

    ' snapshot the current code/designer windows
    Dim keys As New Collection, caps As New Collection
    For Each w In gVBE.Windows
        If w.Visible Then
            If w.Type = vbext_wt_CodeWindow Or w.Type = vbext_wt_Designer Then
                k = NormalizeCaption(w.Caption) & "|" & w.Type
                Err.Clear
                keys.Add k, k
                If Err.Number = 0 Then caps.Add NormalizeCaption(w.Caption), k
            End If
        End If
    Next

    ' MRU-ranked windows first...
    For i = 1 To MRU_Count()
        k = MRU_Key(i)
        Err.Clear
        cap = caps(k)
        If Err.Number = 0 Then
            AddEntry k, cap
            caps.Remove k
            keys.Remove k
        End If
    Next

    ' ...then any windows never activated yet, in collection order
    For i = 1 To keys.Count
        k = keys(i)
        Err.Clear
        cap = caps(k)
        If Err.Number = 0 Then AddEntry k, cap
    Next
End Sub

Private Sub AddEntry(ByVal k As String, ByVal cap As String)
    If lstWins.ListCount > UBound(mItemKeys) Then
        ReDim Preserve mItemKeys(0 To lstWins.ListCount * 2)
    End If
    mItemKeys(lstWins.ListCount) = k
    lstWins.AddItem cap
End Sub

Private Sub SizeAndPosition()
    On Error Resume Next
    Dim cnt As Long
    cnt = lstWins.ListCount
    If cnt > 15 Then cnt = 15
    lstWins.Move 0, 0, 5000, cnt * 255 + 120
    Me.Width = lstWins.Width
    Me.Height = lstWins.Height

    Dim rc As RECT
    GetWindowRect MDIClientHwnd(), rc
    Me.Move ((rc.Left + rc.Right) \ 2) * Screen.TwipsPerPixelX - Me.Width \ 2, _
            ((rc.Top + rc.Bottom) \ 2) * Screen.TwipsPerPixelY - Me.Height \ 2
End Sub

Private Sub ActivateByKey(ByVal k As String)
    On Error Resume Next
    Dim w As VBIDE.Window
    For Each w In gVBE.Windows
        If NormalizeCaption(w.Caption) & "|" & w.Type = k Then
            w.SetFocus
            If gTabBarVisible And w.WindowState <> vbext_ws_Maximize Then
                w.WindowState = vbext_ws_Maximize
            End If
            MRU_Touch k
            Exit Sub
        End If
    Next
End Sub
