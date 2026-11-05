Attribute VB_Name = "modBranch"
Option Explicit

' =====================================================================
'  Branch switching.
'
'  The git half is one command. The problem is everything the IDE is
'  holding: VB6 reads every .frm/.bas/.cls/.ctl into memory when the
'  project opens and never looks at the file again. The moment a
'  checkout rewrites those files the editor is showing the OLD
'  branch's code - silently, with no warning anywhere - and File/Save
'  would then write it straight back over the branch just checked out.
'  The .vbp itself can differ too: a branch may have components the
'  other does not.
'
'  Nothing in the extensibility model reloads a component from disk,
'  so the only correct answer is to make VB6 read the whole project
'  again from scratch:
'
'      VBProjects.AddFromFile(path, Exclusive:=True)
'
'  which is exactly what File/Open Project does - it closes what is
'  open and loads the given .vbp (or .vbg) fresh. Reloading single
'  components instead (VBComponents.Remove + AddFile) was rejected:
'  it rewrites the .vbp's component order, cannot follow a component
'  that exists on only one of the branches, and leaves the project
'  dirty, which is the one state a checkout must never be left in.
'
'  The reload costs the user's open windows, so the add-in puts them
'  back: which components had a code window or a designer open, where
'  the caret and the top line were in each, and which window was
'  active. The tab bar's own left-to-right order survives untouched -
'  frmTabs keys it by caption and keeps it across refreshes.
'
'  Guards, every one of which refuses rather than risk the tree, and
'  each of which says why:
'
'   - the IDE not being free to answer (Mode_Busy): a checkout under
'     a project the IDE is running is never safe. What counts as
'     running is modRunMode's question, not this module's.
'   - any component or project with unsaved IDE changes: the reload
'     throws them away and the user could never get them back.
'   - a project that has never been saved, or several projects that
'     are not a saved project group: there is no single file to reopen
'     the IDE's state from.
'   - tracked changes in the working tree: offered as a stash, or the
'     user can let git carry them across. Untracked files are ignored;
'     git never minds them during a checkout.
'   - the project file missing on the branch just checked out: the
'     reload is skipped and said so loudly, rather than closing the
'     project out from under the user.
'
'  The git commands themselves run quiet (Git_Checkout -> RunGit with
'  pumpMsgs False), so between the checkout starting and the project
'  being reloaded the IDE pumps no messages at all: no Ctrl+S, no
'  timer tick, no add-in poll. gGitSwitching then keeps the tab-bar
'  timer out of the reload itself, which does pump.
' =====================================================================

' One open component window, as it was before the reload.
Private Type WinSnap
    Comp As String
    isDesigner As Boolean
    topLine As Long
    selLine As Long
    selCol As Long
    active As Boolean
End Type

' =====================================================================
'  Entry points
' =====================================================================

' Menu / toolbar / the tab bar's branch label.
Public Sub Branch_ShowPicker()
    On Error Resume Next
    If Mode_Busy("Switching branches") Then Exit Sub
    If Not Git_HasRepo() Then
        MsgBox "The active project is not inside a git repository.", _
               vbInformation, "Switch Branch"
        Exit Sub
    End If
    frmBranches.ShowPicker
End Sub

' Do the switch. Returns True when the working tree actually moved;
' every failure has already been explained to the user by then.
Public Function Branch_SwitchTo(ByVal refName As String, _
        ByVal isRemote As Boolean) As Boolean
    On Error Resume Next

    ' re-checked here, not just when the picker opened: it is modeless,
    ' and the IDE stayed live behind it
    If Mode_Busy("Switching branches") Then Exit Function
    If Not Git_HasRepo() Then Exit Function

    Dim why As String
    If UnsavedWork(why) Then
        MsgBox "These have unsaved changes in the IDE:" & vbCrLf & vbCrLf & _
               why & vbCrLf & _
               "Switching branches reloads the project from disk, which " & _
               "would throw them away." & vbCrLf & _
               "Save them (Ctrl+S) or undo them, then switch.", _
               vbExclamation, "Switch Branch"
        Exit Function
    End If

    Dim problem As String, reopenFile As String
    reopenFile = ReopenPath(problem)
    If Len(reopenFile) = 0 Then
        MsgBox "The branch cannot be switched from here: " & problem & _
               vbCrLf & vbCrLf & _
               "The IDE's copy of the project has to be reloaded from " & _
               "disk after a checkout, and there is no single file to " & _
               "reload it from.", vbExclamation, "Switch Branch"
        Exit Function
    End If

    If Git_HasTrackedChanges() Then
        Select Case MsgBox( _
            "The working tree has uncommitted changes." & vbCrLf & vbCrLf & _
            "Yes" & vbTab & "stash them first (git stash push), then switch" & _
            vbCrLf & _
            "No" & vbTab & "switch and let git carry them over" & vbCrLf & _
            "Cancel" & vbTab & "stay on " & Git_Branch(), _
            vbYesNoCancel Or vbQuestion, "Switch Branch")
        Case vbCancel
            Exit Function
        Case vbYes
            Dim serr As String
            serr = Git_StashTracked("VB6 Modernizr: switching to " & refName)
            If Len(serr) > 0 Then
                MsgBox "git could not stash the changes:" & vbCrLf & vbCrLf & _
                       serr, vbExclamation, "Switch Branch"
                Exit Function
            End If
            MsgBox "Stashed. Bring them back with ""git stash pop"" once " & _
                   "you are on the branch you want them on.", _
                   vbInformation, "Switch Branch"
        End Select
    End If

    ' --- past here the tree moves -------------------------------------
    Dim snaps() As WinSnap, nSnap As Long
    nSnap = SnapWindows(snaps)

    gGitSwitching = True

    Dim gerr As String
    gerr = Git_Checkout(refName, isRemote)
    If Len(gerr) > 0 Then
        ' nothing was reloaded and nothing was closed: the IDE is still
        ' consistent with whatever git left on disk
        Done
        MsgBox "git could not switch to '" & refName & "':" & vbCrLf & _
               vbCrLf & gerr, vbExclamation, "Switch Branch"
        Exit Function
    End If

    Dim loc As String
    loc = Git_LocalNameFor(refName, isRemote)

    If Len(Dir$(reopenFile)) = 0 Then
        Done
        MsgBox "Now on '" & loc & "', but" & vbCrLf & vbCrLf & _
               reopenFile & vbCrLf & vbCrLf & _
               "does not exist on that branch, so the project could not " & _
               "be reloaded." & vbCrLf & vbCrLf & _
               "The IDE is still holding the code of the branch you came " & _
               "from. Do NOT save - close the project, or switch back " & _
               "first.", vbCritical, "Switch Branch"
        Branch_SwitchTo = True
        Exit Function
    End If

    ReloadProject reopenFile, snaps, nSnap
    Done
    Branch_SwitchTo = True
End Function

' Let the pollers back in and bring the add-in's own git state up to
' date: the branch label, the per-tab badges and the changed-line
' marks all read from caches the checkout just invalidated.
Private Sub Done()
    On Error Resume Next
    gGitSwitching = False
    Mode_Invalidate
    Git_RefreshNow
    Git_StatusRefreshSync
    frmTabs.ForceRedraw
    Highlight_InvalidateAll
End Sub

' =====================================================================
'  Guards
' =====================================================================

' Anything a reload from disk would silently discard. Names up to a
' handful of them so the refusal is actionable.
Private Function UnsavedWork(outWhy As String) As Boolean
    On Error Resume Next
    Const MAX_NAMED As Long = 8
    Dim prj As VBIDE.VBProject, Comp As VBIDE.VBComponent, n As Long

    For Each prj In gVBE.VBProjects
        If prj.IsDirty Then
            n = n + 1
            If n <= MAX_NAMED Then _
                outWhy = outWhy & "    " & prj.Name & "  (the project)" & vbCrLf
        End If
        For Each Comp In prj.VBComponents
            If Comp.IsDirty Then
                n = n + 1
                If n <= MAX_NAMED Then _
                    outWhy = outWhy & "    " & Comp.Name & vbCrLf
            End If
        Next
    Next

    If n > MAX_NAMED Then _
        outWhy = outWhy & "    ... and " & (n - MAX_NAMED) & " more" & vbCrLf
    UnsavedWork = (n > 0)
End Function

' The one file File/Open Project would be pointed at to get this state
' back: the project group file when a group is open, otherwise the
' single project's .vbp. "" (with a reason) when the IDE is holding
' something no single file can restore.
Private Function ReopenPath(outProblem As String) As String
    On Error Resume Next

    ' VBProjects.FileName is the group file - set only for a real .vbg
    Dim grp As String
    grp = gVBE.VBProjects.FileName
    If Len(grp) > 0 Then
        If LCase$(Right$(grp, 4)) = ".vbg" Then
            If Len(Dir$(grp)) > 0 Then
                ReopenPath = grp
                Exit Function
            End If
        End If
    End If

    Dim prj As VBIDE.VBProject, f As String, n As Long, one As String
    For Each prj In gVBE.VBProjects
        f = ""
        f = prj.FileName
        If Len(f) = 0 Then
            outProblem = "project '" & prj.Name & "' has never been saved."
            Exit Function
        End If
        n = n + 1
        one = f
    Next

    If n = 1 Then
        ReopenPath = one
    ElseIf n = 0 Then
        outProblem = "no project is open."
    Else
        outProblem = n & " projects are open and they are not a saved " & _
                     "project group (.vbg)."
    End If
End Function

' =====================================================================
'  Reload
' =====================================================================

Private Sub ReloadProject(ByVal path As String, snaps() As WinSnap, _
        ByVal n As Long)
    On Error Resume Next

    ' Exclusive:=True makes this File/Open Project: everything open is
    ' closed and the project is read fresh off the disk git just wrote.
    gVBE.VBProjects.AddFromFile path, True

    ' the open settles over several message passes (designers in
    ' particular), so let the IDE breathe before asking it anything
    Dim i As Long
    For i = 1 To 20
        DoEvents
    Next

    RestoreWindows snaps, n
End Sub

' Which component windows are open, where the caret sits in each, and
' which one is in front. Window captions are the only handle the
' object model gives on a window's component, and they are what
' frmTabs keys on too: "Project1 - Form1 (Code)" -> "Form1".
Private Function SnapWindows(snaps() As WinSnap) As Long
    On Error Resume Next
    Dim n As Long, w As VBIDE.Window, aw As VBIDE.Window
    Dim actKey As String, nm As String
    ReDim snaps(0 To 16)

    Set aw = gVBE.ActiveWindow
    If Not aw Is Nothing Then actKey = WinKey(aw)

    For Each w In gVBE.Windows
        If w.Visible And TabbableWindow(w) Then
            nm = CompNameOf(w)
            If Len(nm) > 0 Then
                If n > UBound(snaps) Then ReDim Preserve snaps(0 To n * 2)
                snaps(n).Comp = nm
                snaps(n).isDesigner = (w.Type = vbext_wt_Designer)
                snaps(n).active = (Len(actKey) > 0 And WinKey(w) = actKey)
                If Not snaps(n).isDesigner Then CaretOf nm, snaps(n)
                n = n + 1
            End If
        End If
    Next
    SnapWindows = n
End Function

Private Sub RestoreWindows(snaps() As WinSnap, ByVal n As Long)
    On Error Resume Next
    Dim i As Long, Comp As VBIDE.VBComponent, w As VBIDE.Window
    Dim actW As VBIDE.Window

    For i = 0 To n - 1
        Set Comp = FindComp(snaps(i).Comp)
        ' gone on this branch: nothing to reopen, and nothing is wrong
        If Not Comp Is Nothing Then
            Set w = Nothing
            If snaps(i).isDesigner Then
                Err.Clear
                Set w = Comp.DesignerWindow
                If Err.Number <> 0 Then Set w = Nothing
                If Not w Is Nothing Then w.Visible = True
            Else
                Set w = ShowCodeWindow(Comp, snaps(i))
            End If
            If Not w Is Nothing Then
                If gTabBarVisible And w.WindowState <> vbext_ws_Maximize Then _
                    w.WindowState = vbext_ws_Maximize
                If snaps(i).active Then Set actW = w
            End If
        End If
        DoEvents
    Next

    If Not actW Is Nothing Then
        ActivateIDEWindow actW
        actW.SetFocus
    End If
End Sub

' Open a component's code window and put the caret and the scroll
' position back. Both are clamped: the branch just checked out may
' have a far shorter module than the one the caret was recorded in.
Private Function ShowCodeWindow(ByVal Comp As VBIDE.VBComponent, _
        sn As WinSnap) As VBIDE.Window
    On Error Resume Next
    Dim cm As VBIDE.CodeModule, cp As VBIDE.CodePane, ln As Long
    Set cm = Comp.CodeModule
    If cm Is Nothing Then Exit Function

    Set cp = cm.CodePane
    If cp Is Nothing Then Exit Function
    cp.Show
    Set ShowCodeWindow = cp.Window

    ln = sn.selLine
    If ln > cm.CountOfLines Then ln = cm.CountOfLines
    If ln >= 1 Then
        Dim col As Long
        col = sn.selCol
        If col < 1 Then col = 1
        cp.SetSelection ln, col, ln, col
    End If

    ln = sn.topLine
    If ln > cm.CountOfLines Then ln = cm.CountOfLines
    If ln >= 1 Then cp.TopLine = ln
End Function

Private Sub CaretOf(ByVal compName As String, sn As WinSnap)
    On Error Resume Next
    Dim Comp As VBIDE.VBComponent, cp As VBIDE.CodePane
    Dim sl As Long, sc As Long, el As Long, ec As Long

    Set Comp = FindComp(compName)
    If Comp Is Nothing Then Exit Sub
    Set cp = Comp.CodeModule.CodePane
    If cp Is Nothing Then Exit Sub

    cp.GetSelection sl, sc, el, ec
    sn.selLine = sl
    sn.selCol = sc
    sn.topLine = cp.TopLine
End Sub

Private Function WinKey(ByVal w As VBIDE.Window) As String
    On Error Resume Next
    WinKey = NormalizeCaption(w.Caption) & "|" & w.Type
End Function

' Component name behind a window, or "" for a window that has none
' (the Object Browser, which is tabbable but has no component).
Private Function CompNameOf(ByVal w As VBIDE.Window) As String
    On Error Resume Next
    If w.Type = vbext_wt_Browser Then Exit Function
    Dim nm As String, p As Long
    nm = NormalizeCaption(w.Caption)
    p = InStrRev(nm, " (")
    If p > 0 Then nm = Left$(nm, p - 1)
    CompNameOf = nm
End Function

Private Function FindComp(ByVal nm As String) As VBIDE.VBComponent
    On Error Resume Next
    Dim prj As VBIDE.VBProject, c As VBIDE.VBComponent
    For Each prj In gVBE.VBProjects
        Set c = Nothing
        Err.Clear
        Set c = prj.VBComponents(nm)
        If Err.Number = 0 Then
            If Not c Is Nothing Then
                Set FindComp = c
                Exit Function
            End If
        End If
    Next
End Function
