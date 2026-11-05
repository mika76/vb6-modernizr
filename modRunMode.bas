Attribute VB_Name = "modRunMode"
Option Explicit

' =====================================================================
'  Design / run / break gating.
'
'  The IDE runs a project IN PROCESS, on the IDE's own thread, so when
'  the user hits F5 nothing about the add-in goes quiet by itself. But
'  "the IDE is not in design mode" covers two very different
'  situations, and only one of them is dangerous:
'
'   - [run]: project code owns the thread. Walking the VBE object model
'     from a timer tick re-enters the IDE while it is executing that
'     code, which is exactly what breaks it. The add-in stands down.
'
'   - [break]: a breakpoint (or Step Into) has suspended the project
'     and the IDE is back to pumping its own message loop, waiting for
'     the user. Code panes are live, the object model answers, and
'     stepping through code is when line numbers, the tab bar and the
'     editor shortcuts are wanted most. So the add-in stays awake:
'     Mode_IsIdle() is true for design and break alike, and it is what
'     nearly everything asks.
'
'  Keys are not gated here at all. Every shortcut in modWheel already
'  requires focus to be in an IDE code pane, or in one of the add-in's
'  own text boxes; a running program's windows are unowned top-level
'  windows that satisfy neither. A blanket gate took Ctrl+F away from
'  the IDE's own editor without protecting the running program from
'  anything it was not already safe from.
'
'  Detection comes in two steps, because neither source answers both
'  questions:
'
'  1. Design or not, from the IDE's title bar, which always carries the
'     state in the first brackets - "Project1 - Microsoft Visual Basic
'     [design]", "[run]", "[break]". Reading it is a plain
'     GetWindowText, safe from inside a message hook or a paint
'     handler. The design word is baselined at connect (always design
'     mode) rather than compared against "design", so a localized IDE
'     works too.
'
'  2. Break or run, from VBProject.Mode. The title cannot answer this
'     one: the design word is known, but not which of the other two is
'     on screen. VBProject.Mode cannot answer step 1 - VB6 leaves it at
'     vbext_vm_Run for a project sitting in the designer, so gating on
'     it switched the whole add-in off - but the test here is the
'     positive one, = vbext_vm_Break, which that quirk never produces.
'     A read that fails or lies therefore costs at most a stood-down
'     break mode, never a stood-down design mode.
'
'     The title word in place when Mode first says Break IS the break
'     word, so it is remembered. After that break is recognized from
'     the title alone, and the object model is left out of it.
' =====================================================================

Private Const ST_DESIGN As Long = 0
Private Const ST_BREAK As Long = 1
Private Const ST_RUN As Long = 2

Private Const CACHE_SEC As Single = 0.15

Private mStamp As Single
Private mKnown As Boolean
Private mState As Long

Private mDesignTag As String      ' the title's state word in design mode
Private mBreakTag As String       ' ... and in break mode, once learned
Private mMainHwnd As Long         ' cached: the IDE frame outlives us
Private mWasIdle As Boolean

' ---------------------------------------------------------------------

' Called once from InitAddin, in design mode by definition.
Public Sub Mode_Init()
    On Error Resume Next
    mMainHwnd = MainHwnd()
    mDesignTag = StateTag()
    mWasIdle = True
End Sub

' True in design mode only.
Public Function Mode_IsDesign() As Boolean
    Mode_IsDesign = (State() = ST_DESIGN)
End Function

' True whenever the IDE is pumping its own message loop and the object
' model answers: design mode, and a project suspended at a breakpoint.
' Reads the title and nothing else, so it is safe to ask from a message
' hook or a paint handler.
Public Function Mode_IsIdle() As Boolean
    Mode_IsIdle = (State() <> ST_RUN)
End Function

' Make the next ask go to the title bar again.
Public Sub Mode_Invalidate()
    mKnown = False
End Sub

' Called from the tab-bar timer: keeps the cached answer fresh even
' when no input is reaching the hook, learns the break word the first
' time a project stops at a breakpoint, and repaints what went stale
' once the IDE is interactive again - nothing drew the tab bar or the
' editor overlays while project code owned the thread, and no paint is
' coming for them on its own.
Public Sub Mode_Poll()
    On Error Resume Next
    Mode_Invalidate

    Dim learned As Boolean
    If State() = ST_RUN Then
        learned = LearnBreakTag()
        If learned Then Mode_Invalidate
    End If

    ' "learned" repaints too, and happens once per IDE session at most:
    ' a project stopped at a breakpoint counted as running until the
    ' word was known, so up to one tick's worth of overlays was skipped.
    Dim idle As Boolean
    idle = Mode_IsIdle()
    If idle = mWasIdle And Not learned Then Exit Sub
    mWasIdle = idle
    If idle Then
        frmTabs.ForceRedraw
        Highlight_InvalidateAll
    End If
End Sub

' For user-invoked commands: true (having said so) when project code is
' running and the command needs the object model. A project stopped at
' a breakpoint is not busy - the IDE is interactive there, and so is
' the add-in.
Public Function Mode_Busy(ByVal what As String) As Boolean
    On Error Resume Next
    Mode_Invalidate
    ' a menu command can afford the object-model check a hook cannot
    If State() = ST_RUN Then
        If LearnBreakTag() Then Mode_Invalidate
    End If
    If Mode_IsIdle() Then Exit Function

    Mode_Busy = True
    MsgBox what & " is not available while the IDE is running a " & _
           "project." & vbCrLf & vbCrLf & _
           "Stop the program (Run" & Chr$(149) & "End), or break into " & _
           "it, and try again.", _
           vbInformation, "VB6 Modernizr"
End Function

' ---------------------------------------------------------------------

Private Function State() As Long
    Dim t As Single
    t = Timer
    ' t < mStamp only at the midnight wrap - just re-ask then
    If mKnown And t >= mStamp And (t - mStamp) < CACHE_SEC Then
        State = mState
        Exit Function
    End If
    mState = QueryState()
    mStamp = t
    mKnown = True
    State = mState
End Function

Private Function QueryState() As Long
    On Error Resume Next
    Dim t As String
    t = StateTag()

    ' No title to read (too early, or the frame went away): the add-in
    ' stays useful rather than switching itself off.
    If Len(t) = 0 Then Exit Function

    ' First word we ever see is the design word - Mode_Init normally
    ' catches it, this covers a connect that beat the title into place.
    If Len(mDesignTag) = 0 Then mDesignTag = t

    If t = mDesignTag Then Exit Function              ' ST_DESIGN
    If Len(mBreakTag) > 0 And t = mBreakTag Then
        QueryState = ST_BREAK
        Exit Function
    End If
    QueryState = ST_RUN
End Function

' The one object-model call in this module, and the one place it is made
' from: a timer tick or a menu command, never a message hook and never a
' paint handler. Returns True when it has just learned the break word.
Private Function LearnBreakTag() As Boolean
    On Error Resume Next
    If Len(mBreakTag) > 0 Then Exit Function
    If gVBE Is Nothing Then Exit Function

    Dim before As String, after As String, broke As Boolean
    before = StateTag()
    If Len(before) = 0 Or before = mDesignTag Then Exit Function

    Dim prj As VBIDE.VBProject
    For Each prj In gVBE.VBProjects
        If prj.Mode = vbext_vm_Break Then
            broke = True
            Exit For
        End If
    Next
    If Not broke Then Exit Function

    ' The title has to still say what it said before the read, or we
    ' caught a transition and would pin the wrong word for good.
    after = StateTag()
    If after <> before Then Exit Function

    mBreakTag = after
    LearnBreakTag = True
End Function

' The state word between the first pair of brackets in the IDE title,
' lower-cased. "" when the title has none.
Private Function StateTag() As String
    On Error Resume Next
    If IsWindow(mMainHwnd) = 0 Then mMainHwnd = MainHwnd()
    If mMainHwnd = 0 Then Exit Function

    Dim s As String, p As Long, q As Long
    s = WndText(mMainHwnd)
    p = InStr(s, "[")
    If p = 0 Then Exit Function
    q = InStr(p + 1, s, "]")
    If q = 0 Then Exit Function
    StateTag = LCase$(Trim$(Mid$(s, p + 1, q - p - 1)))
End Function
