Attribute VB_Name = "modNav"
Option Explicit

' =====================================================================
'  Back / Forward navigation history, VS-style, session only.
'
'  The tab-bar timer feeds Nav_Poll with the caret position. Small
'  movements just update the current entry; switching modules or
'  jumping more than NAV_NEAR lines pushes a new entry and drops any
'  forward tail. Nav_Back / Nav_Forward walk the list; landing on an
'  entry leaves the caret exactly on it, so the next poll reads as
'  drift and nothing spurious is recorded.
' =====================================================================

Private Const NAV_MAX As Long = 64
Private Const NAV_NEAR As Long = 10   ' lines: drift vs. a real jump

Private Type NavEntry
    Proj As String
    Comp As String
    Line As Long
End Type

Private mHist() As NavEntry
Private mCount As Long
Private mCur As Long          ' index of the entry the caret is "on"
Private mBackBtns As Collection   ' all Back buttons (menu + toolbar),
Private mFwdBtns As Collection    ' kept Enabled-synced with the list

Public Sub Nav_RegisterButtons(btnBack As Object, btnFwd As Object)
    On Error Resume Next
    If mBackBtns Is Nothing Then
        Set mBackBtns = New Collection
        Set mFwdBtns = New Collection
    End If
    mBackBtns.Add btnBack
    mFwdBtns.Add btnFwd
    SyncButtons
End Sub

Public Sub Nav_Term()
    mCount = 0
    mCur = 0
    Set mBackBtns = Nothing
    Set mFwdBtns = Nothing
End Sub

Public Sub Nav_Poll()
    On Error Resume Next
    Dim cp As VBIDE.CodePane
    Set cp = gVBE.ActiveCodePane
    If cp Is Nothing Then Exit Sub

    Dim sl As Long, sc As Long, el As Long, ec As Long
    cp.GetSelection sl, sc, el, ec
    If sl < 1 Then Exit Sub

    Dim comp As String, proj As String
    comp = cp.CodeModule.Parent.Name
    proj = cp.CodeModule.Parent.Collection.Parent.Name
    If Len(comp) = 0 Then Exit Sub

    If mCount = 0 Then
        ReDim mHist(0 To NAV_MAX - 1)
        mHist(0).Proj = proj
        mHist(0).Comp = comp
        mHist(0).Line = sl
        mCount = 1
        mCur = 0
        SyncButtons
        Exit Sub
    End If

    ' drift: follow the caret, don't record
    If mHist(mCur).Comp = comp And mHist(mCur).Proj = proj And _
       Abs(mHist(mCur).Line - sl) <= NAV_NEAR Then
        mHist(mCur).Line = sl
        Exit Sub
    End If

    ' real jump: drop the forward tail, append
    mCount = mCur + 1
    If mCount = NAV_MAX Then           ' full: forget the oldest
        Dim i As Long
        For i = 0 To NAV_MAX - 2
            mHist(i) = mHist(i + 1)
        Next
        mCount = NAV_MAX - 1
    End If
    mHist(mCount).Proj = proj
    mHist(mCount).Comp = comp
    mHist(mCount).Line = sl
    mCur = mCount
    mCount = mCount + 1
    SyncButtons
End Sub

Public Sub Nav_Back()
    If mCur <= 0 Then Exit Sub
    mCur = mCur - 1
    JumpTo mCur
End Sub

Public Sub Nav_Forward()
    If mCur >= mCount - 1 Then Exit Sub
    mCur = mCur + 1
    JumpTo mCur
End Sub

Private Sub JumpTo(ByVal i As Long)
    On Error Resume Next
    Dim m As MatchInfo
    m.Proj = mHist(i).Proj
    m.Comp = mHist(i).Comp
    m.LineNum = mHist(i).Line
    m.Col = 1
    m.MatchLen = 0
    GoToMatch m
    SyncButtons
End Sub

Private Sub SyncButtons()
    On Error Resume Next
    If mBackBtns Is Nothing Then Exit Sub
    Dim i As Long
    For i = 1 To mBackBtns.Count
        mBackBtns(i).Enabled = (mCur > 0)
    Next
    For i = 1 To mFwdBtns.Count
        mFwdBtns(i).Enabled = (mCur < mCount - 1)
    Next
End Sub
