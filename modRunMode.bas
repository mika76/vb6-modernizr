Attribute VB_Name = "modRunMode"
Option Explicit

' =====================================================================
'  Design / run gating.
'
'  The IDE runs a project IN PROCESS, on the IDE's own thread, so when
'  the user hits F5 nothing about the add-in goes quiet by itself: the
'  400 ms tab-bar timer keeps firing inside the running program's
'  message loop, and every poll behind it (tab bar, gutter, bookmarks,
'  git, navigation, auto-backup) either walks the VBE object model or
'  shells out. Re-entering the IDE while it is busy with project code
'  is what breaks it, so everything automatic asks Mode_IsDesign()
'  first, and menu commands say so and decline.
'
'  Keys are NOT gated here. Every shortcut in modWheel already requires
'  focus to be in an IDE code pane, in one of the add-in's own text
'  boxes, or over a window the add-in subclassed, and a running
'  program's windows are unowned top-level windows that are none of
'  those - that is what keeps the add-in's hands off the program.
'  Gating the hook on the mode as well only took Ctrl+F, Ctrl+D and the
'  rest away from the IDE's own editor for as long as a program was
'  loaded, which is a thing users noticed and complained about.
'
'  BREAK MODE IS NOT TREATED AS DESIGN MODE, and not for want of
'  trying. The reasoning that says it should is sound as far as it
'  goes: at a breakpoint the project is suspended, the IDE is pumping
'  its own message loop again, and the object model answers - verified
'  against D:\vb6bench spike-make, where a second add-in enumerated
'  VBComponents and FileNames in [break] without complaint. Stepping
'  through code is also when the tab bar and the line numbers are
'  wanted most, and the gutter in particular leaves numbers behind from
'  before the step.
'
'  But letting this add-in's polls run at a breakpoint hangs the IDE.
'  Measured: 0% CPU and responsive in [design]; 100% of one core and
'  not responding within a second of reaching [break], every time. It
'  is NOT the editor overlays - the same hang happens with the line
'  numbers and the indentation guides both switched off, which takes
'  Highlight_Active() to False and makes the gutter and the highlight
'  paint paths no-ops. So it is something in the timer's own poll set
'  (RefreshTabs, BM_Poll, Nav_Poll, Backup_Poll) looping against the
'  IDE while it is in break mode; RefreshTabs forcing the active MDI
'  child back to maximized while the IDE is re-docking the Immediate
'  window is the obvious suspect, but it has not been pinned down.
'  Until it is, break mode stands down with run mode. Do not relax this
'  without reproducing that measurement.
'
'  The state comes from the IDE's own title bar, which always carries
'  it in the first brackets - "Project1 - Microsoft Visual Basic
'  [design]", "[run]", "[break]" (confirmed for all three). A maximized
'  MDI child appends a second pair, as in "... [design] - [frmMain
'  (Code)]", so only the FIRST pair is read. VBProject.Mode looks like
'  the proper answer and is not: VB6 leaves it at vbext_vm_Run for a
'  project sitting in the designer, so gating on it switched the whole
'  add-in off. Reading the title is a plain GetWindowText - no COM call
'  from inside a message hook or a paint handler.
'
'  The word itself is localized, so it is not compared against
'  "design": the add-in connects at IDE startup, which is always
'  design mode, so whatever word is in the title then IS the design
'  word. Everything else means the IDE is busy with project code.
' =====================================================================

Private Const CACHE_SEC As Single = 0.15

Private mStamp As Single
Private mKnown As Boolean
Private mIsDesign As Boolean
Private mWasDesign As Boolean

Private mDesignTag As String      ' the title's state word in design mode
Private mMainHwnd As Long         ' cached: the IDE frame outlives us

' ---------------------------------------------------------------------

' Called once from InitAddin, in design mode by definition.
Public Sub Mode_Init()
    On Error Resume Next
    mMainHwnd = MainHwnd()
    mDesignTag = StateTag()
    mWasDesign = True
End Sub

Public Function Mode_IsDesign() As Boolean
    Dim t As Single
    t = Timer
    ' t < mStamp only at the midnight wrap - just re-ask then
    If mKnown And t >= mStamp And (t - mStamp) < CACHE_SEC Then
        Mode_IsDesign = mIsDesign
        Exit Function
    End If
    mIsDesign = QueryDesign()
    mStamp = t
    mKnown = True
    Mode_IsDesign = mIsDesign
End Function

' Make the next ask go to the title bar again.
Public Sub Mode_Invalidate()
    mKnown = False
End Sub

' Called from the tab-bar timer: keeps the cached answer fresh even
' when no input is reaching the hook, and repaints what went stale once
' a run ends. Nothing drew the tab bar or the editor overlays while
' project code owned the thread, and no paint is coming for them on
' their own - the tab bar especially, because its change signature
' still matches, so a plain refresh would decide there was nothing to
' do and leave whatever Form_Resize last baked into the AutoRedraw
' bitmap on screen.
Public Sub Mode_Poll()
    On Error Resume Next
    Mode_Invalidate
    Dim d As Boolean
    d = Mode_IsDesign()
    If d = mWasDesign Then Exit Sub
    mWasDesign = d
    If d Then
        frmTabs.ForceRedraw
        Highlight_InvalidateAll
    End If
End Sub

' For user-invoked commands: true (having said so) when the IDE is
' running or debugging a project and the command needs the object
' model.
Public Function Mode_Busy(ByVal what As String) As Boolean
    On Error Resume Next
    Mode_Invalidate
    If Mode_IsDesign() Then Exit Function
    Mode_Busy = True
    MsgBox what & " is not available while the IDE is running or " & _
           "debugging a project." & vbCrLf & vbCrLf & _
           "Stop the program (Run" & Chr$(149) & "End) and try again.", _
           vbInformation, "VB6 Modernizr"
End Function

' ---------------------------------------------------------------------

Private Function QueryDesign() As Boolean
    On Error Resume Next
    Dim t As String
    t = StateTag()

    ' No title to read (too early, or the frame went away): the add-in
    ' stays useful rather than switching itself off.
    If Len(t) = 0 Then QueryDesign = True: Exit Function

    ' First word we ever see is the design word - Mode_Init normally
    ' catches it, this covers a connect that beat the title into place.
    If Len(mDesignTag) = 0 Then mDesignTag = t

    QueryDesign = (t = mDesignTag)
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
