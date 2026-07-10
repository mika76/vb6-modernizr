VERSION 5.00
Begin VB.Form frmQuickOpen 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   5100
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6600
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   340
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   440
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFilter 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6360
   End
   Begin VB6Modernizr.ucList lstResults 
      Height          =   4200
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   7408
   End
End
Attribute VB_Name = "frmQuickOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  Quick Open (Ctrl+P): VS Code style go-to-file palette. Type to
'  fuzzy-filter every component of the open project group by module
'  or file name; capital-letter humps count as word starts, so "ms"
'  finds modSubclass and "ff" finds frmFindFiles. Enter opens the code
'  window (Shift+Enter the designer), Esc or clicking away cancels.
'  An empty filter lists MRU windows first, then the rest A-Z.
' =====================================================================

Private Const VIS_ROWS As Long = 14
Private Const FOOT_H As Long = 18      ' px at 96 dpi, legend footer

' snapshot of all components, taken each time the palette opens
Private mNames() As String
Private mFiles() As String     ' file name tail, "" if never saved
Private mProjs() As String
Private mTotal As Long
Private mMulti As Boolean      ' more than one project open
Private mMRU As Collection     ' MRU rank keyed by "k" & LCase$(name)

' filtered rows -> snapshot index (parallel to lstResults)
Private mIdx() As Long
Private mResCount As Long

Public Sub ShowQuickOpen()
    On Error Resume Next
    Load Me
    SetWindowLongA Me.hwnd, GWL_HWNDPARENT, MainHwnd()
    Snapshot
    SizeAndPosition
    If Len(txtFilter.Text) > 0 Then
        txtFilter.Text = ""            ' fires Change -> RebuildList
    Else
        RebuildList
    End If
    Me.Show vbModeless
    txtFilter.SetFocus
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    Me.Hide
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim w As Long, h As Long, footH As Long
    w = Me.ScaleWidth
    h = Me.ScaleHeight
    footH = ScaleForDpi(FOOT_H)

    Me.Cls
    ' footer band with the key legend
    Me.Line (1, h - footH - 1)-(w - 2, h - 2), vbWindowBackground, BF
    Me.ForeColor = vbGrayText
    Me.CurrentX = ScaleForDpi(8)
    Me.CurrentY = h - footH + (footH - Me.TextHeight("X")) \ 2 - 1
    Me.Print "Enter = open code    Shift+Enter = open form    Esc = close"
    Me.Line (0, 0)-(w - 1, h - 1), THEME_BORDER, B
End Sub

' ---------------------------------------------------------------------
'  Input
' ---------------------------------------------------------------------

Private Sub txtFilter_Change()
    RebuildList
End Sub

Private Sub txtFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Dim i As Long
    Select Case KeyCode
    Case vbKeyDown, vbKeyUp, vbKeyPageDown, vbKeyPageUp
        i = lstResults.ListIndex
        Select Case KeyCode
            Case vbKeyDown:     i = i + 1
            Case vbKeyUp:       i = i - 1
            Case vbKeyPageDown: i = i + VIS_ROWS - 1
            Case vbKeyPageUp:   i = i - (VIS_ROWS - 1)
        End Select
        If i < 0 Then i = 0
        If i >= lstResults.ListCount Then i = lstResults.ListCount - 1
        If lstResults.ListCount > 0 Then lstResults.ListIndex = i
        KeyCode = 0
    Case vbKeyReturn
        KeyCode = 0
        OpenSelection (Shift And vbShiftMask) <> 0
    Case vbKeyEscape
        KeyCode = 0
        Me.Hide
    End Select
End Sub

' swallow the ding for keys KeyDown already handled
Private Sub txtFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 27 Then KeyAscii = 0
End Sub

Private Sub lstResults_Click()
    On Error Resume Next
    txtFilter.SetFocus             ' click selects, typing continues
End Sub

Private Sub lstResults_DblClick()
    OpenSelection False
End Sub

Private Sub lstResults_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        OpenSelection (Shift And vbShiftMask) <> 0
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0
        Me.Hide
    End If
End Sub

' ---------------------------------------------------------------------
'  Component snapshot / filtering
' ---------------------------------------------------------------------

Private Sub Snapshot()
    On Error Resume Next
    Dim proj As VBIDE.VBProject, comp As VBIDE.VBComponent
    Dim f As String, i As Long, p As Long, k As String

    mTotal = 0
    ReDim mNames(0 To 128)
    ReDim mFiles(0 To 128)
    ReDim mProjs(0 To 128)
    mMulti = (gVBE.VBProjects.Count > 1)

    For Each proj In gVBE.VBProjects
        For Each comp In proj.VBComponents
            If mTotal > UBound(mNames) Then
                ReDim Preserve mNames(0 To mTotal * 2)
                ReDim Preserve mFiles(0 To mTotal * 2)
                ReDim Preserve mProjs(0 To mTotal * 2)
            End If
            f = ""
            Err.Clear
            f = comp.FileNames(1)
            If Err.Number <> 0 Then f = ""
            p = InStrRev(f, "\")
            If p > 0 Then f = Mid$(f, p + 1)
            mNames(mTotal) = comp.Name
            mFiles(mTotal) = f
            mProjs(mTotal) = proj.Name
            mTotal = mTotal + 1
        Next
    Next

    ' MRU rank by component name, for boosting recently used entries
    Set mMRU = New Collection
    For i = 1 To MRU_Count()
        k = MRU_Key(i)
        p = InStr(k, "|")
        If p > 0 Then k = Left$(k, p - 1)
        p = InStrRev(k, " (")
        If p > 0 Then k = Left$(k, p - 1)
        Err.Clear
        mMRU.Add i, "k" & LCase$(Trim$(k))
    Next
End Sub

Private Sub RebuildList()
    On Error Resume Next
    Dim q As String
    q = Trim$(txtFilter.Text)

    Dim sc() As Long, ord() As Long, n As Long
    ReDim sc(0 To mTotal + 1)
    ReDim ord(0 To mTotal + 1)

    Dim i As Long, s As Long, sf As Long
    For i = 0 To mTotal - 1
        s = FuzzyScore(q, mNames(i))
        sf = FuzzyScore(q, mFiles(i))
        If sf - 5 > s Then s = sf - 5      ' name hits outrank file hits
        If s >= 0 Then
            sc(n) = s * 100 + MRUBoost(mNames(i))
            ord(n) = i
            n = n + 1
        End If
    Next

    SortByScore sc, ord, n

    lstResults.Clear
    mResCount = 0
    ReDim mIdx(0 To n + 1)

    Dim disp As String, icf As String
    For i = 0 To n - 1
        disp = mFiles(ord(i))
        If Len(disp) = 0 Then disp = mNames(ord(i))
        If mMulti Then disp = disp & "  (" & mProjs(ord(i)) & ")"
        icf = mFiles(ord(i))
        If Len(icf) = 0 Then icf = FileForComponent(mNames(ord(i)))
        lstResults.AddItem disp, , icf
        mIdx(mResCount) = ord(i)
        mResCount = mResCount + 1
    Next
    If mResCount > 0 Then lstResults.ListIndex = 0
End Sub

' insertion sort: score desc, ties by name asc (n is small)
Private Sub SortByScore(sc() As Long, ord() As Long, ByVal n As Long)
    Dim i As Long, j As Long, ts As Long, ti As Long
    For i = 1 To n - 1
        ts = sc(i)
        ti = ord(i)
        j = i - 1
        Do While j >= 0
            If sc(j) > ts Then Exit Do
            If sc(j) = ts Then
                If StrComp(mNames(ord(j)), mNames(ti), vbTextCompare) <= 0 _
                    Then Exit Do
            End If
            sc(j + 1) = sc(j)
            ord(j + 1) = ord(j)
            j = j - 1
        Loop
        sc(j + 1) = ts
        ord(j + 1) = ti
    Next
End Sub

Private Function MRUBoost(ByVal nm As String) As Long
    On Error Resume Next
    If mMRU Is Nothing Then Exit Function
    Dim r As Long
    Err.Clear
    r = mMRU("k" & LCase$(nm))
    If Err.Number = 0 And r > 0 And r < 99 Then MRUBoost = 99 - r
End Function

' ---------------------------------------------------------------------
'  Fuzzy matching. -1 = no match; otherwise best of three strategies:
'  whole substring, word-boundary humps ("ms" -> modSubclass), plain
'  subsequence. Empty query matches everything at score 0.
' ---------------------------------------------------------------------

Private Function FuzzyScore(ByVal q As String, ByVal s As String) As Long
    Dim best As Long, sc As Long
    If Len(q) = 0 Then Exit Function
    best = -1

    sc = SubstrScore(q, s)
    If sc > best Then best = sc
    sc = HumpScore(q, s)
    If sc > best Then best = sc
    sc = SubseqScore(q, s)
    If sc > best Then best = sc

    FuzzyScore = best
End Function

Private Function SubstrScore(ByVal q As String, ByVal s As String) As Long
    Dim p As Long, sc As Long
    SubstrScore = -1
    p = InStr(1, s, q, vbTextCompare)
    If p = 0 Then Exit Function
    sc = 300 - p * 2
    If p = 1 Then sc = sc + 200                       ' prefix
    If p = 1 And Len(q) = Len(s) Then sc = sc + 100   ' exact
    If sc < 1 Then sc = 1
    SubstrScore = sc
End Function

' every query char consumed at a word start: first char, an uppercase
' hump, a char after a non-alphanumeric, or a first digit
Private Function HumpScore(ByVal q As String, ByVal s As String) As Long
    Dim qi As Long, i As Long, firstHit As Long, sc As Long
    HumpScore = -1
    qi = 1
    For i = 1 To Len(s)
        If IsBoundary(s, i) Then
            If StrComp(Mid$(s, i, 1), Mid$(q, qi, 1), vbTextCompare) = 0 Then
                If qi = 1 Then firstHit = i
                qi = qi + 1
                If qi > Len(q) Then
                    sc = 350 - firstHit
                    If sc < 1 Then sc = 1
                    HumpScore = sc
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Private Function SubseqScore(ByVal q As String, ByVal s As String) As Long
    Dim qi As Long, i As Long, sc As Long
    Dim lastHit As Long, firstHit As Long
    SubseqScore = -1
    qi = 1
    For i = 1 To Len(s)
        If StrComp(Mid$(s, i, 1), Mid$(q, qi, 1), vbTextCompare) = 0 Then
            If qi = 1 Then firstHit = i
            sc = sc + 1
            If i = lastHit + 1 And lastHit > 0 Then sc = sc + 2
            If IsBoundary(s, i) Then sc = sc + 4
            lastHit = i
            qi = qi + 1
            If qi > Len(q) Then
                sc = 100 + sc - firstHit
                If sc < 1 Then sc = 1
                SubseqScore = sc
                Exit Function
            End If
        End If
    Next
End Function

Private Function IsBoundary(ByVal s As String, ByVal i As Long) As Boolean
    Dim c As String, p As String
    If i = 1 Then IsBoundary = True: Exit Function
    c = Mid$(s, i, 1)
    p = Mid$(s, i - 1, 1)
    If Not IsAlnum(p) Then IsBoundary = True: Exit Function
    If c >= "A" And c <= "Z" Then
        IsBoundary = Not (p >= "A" And p <= "Z")
    ElseIf c >= "0" And c <= "9" Then
        IsBoundary = Not (p >= "0" And p <= "9")
    End If
End Function

Private Function IsAlnum(ByVal c As String) As Boolean
    IsAlnum = (c >= "0" And c <= "9") Or (c >= "A" And c <= "Z") _
           Or (c >= "a" And c <= "z")
End Function

' ---------------------------------------------------------------------
'  Layout / open
' ---------------------------------------------------------------------

Private Sub SizeAndPosition()
    On Error Resume Next
    Dim pad As Long, w As Long, h As Long, listH As Long, y As Long
    pad = ScaleForDpi(8)
    w = ScaleForDpi(440)
    listH = VIS_ROWS * ScaleForDpi(20)

    txtFilter.Move pad, pad, w - pad * 2
    y = pad + txtFilter.Height + ScaleForDpi(6)
    lstResults.Move pad, y, w - pad * 2, listH
    h = y + listH + ScaleForDpi(4) + ScaleForDpi(FOOT_H) + 2

    ' centered near the top of the MDI area, like an editor palette
    Dim rc As RECT
    GetWindowRect MDIClientHwnd(), rc
    Me.Move ((rc.Left + rc.Right) \ 2 - w \ 2) * Screen.TwipsPerPixelX, _
            (rc.Top + ScaleForDpi(40)) * Screen.TwipsPerPixelY, _
            w * Screen.TwipsPerPixelX, h * Screen.TwipsPerPixelY
End Sub

Private Sub OpenSelection(ByVal wantDesigner As Boolean)
    On Error Resume Next
    Dim i As Long
    i = lstResults.ListIndex
    If i < 0 And mResCount > 0 Then i = 0
    If i < 0 Or i >= mResCount Then Exit Sub

    Dim nm As String, pj As String
    nm = mNames(mIdx(i))
    pj = mProjs(mIdx(i))
    Me.Hide

    Dim proj As VBIDE.VBProject, comp As VBIDE.VBComponent
    For Each proj In gVBE.VBProjects
        If StrComp(proj.Name, pj, vbTextCompare) = 0 Then
            Err.Clear
            Set comp = proj.VBComponents(nm)
            If Err.Number <> 0 Then Set comp = Nothing
            Exit For
        End If
    Next
    If comp Is Nothing Then Exit Sub

    Dim w As VBIDE.Window
    If wantDesigner Then
        Err.Clear
        Set w = comp.DesignerWindow
        If Err.Number <> 0 Then Set w = Nothing
    End If
    If w Is Nothing Then
        comp.CodeModule.CodePane.Show
        Set w = comp.CodeModule.CodePane.Window
    End If
    w.Visible = True
    w.SetFocus
    If gTabBarVisible And w.WindowState <> vbext_ws_Maximize Then
        w.WindowState = vbext_ws_Maximize
    End If
    MRU_Touch NormalizeCaption(w.Caption) & "|" & w.Type
End Sub
