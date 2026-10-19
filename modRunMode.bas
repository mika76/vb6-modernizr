Attribute VB_Name = "modRunMode"
Option Explicit

' =====================================================================
'  Design / run / break gating.
'
'  The IDE runs a project IN PROCESS, on the IDE's own thread. Nothing
'  about the add-in goes quiet when the user hits F5, so without a gate:
'
'   - the WH_GETMESSAGE hook in modWheel sees every message of the
'     program being debugged. Keys the add-in claims (F2/F3/F12,
'     Ctrl+P, Ctrl+D, Alt+arrows, the clipboard set) are taken away
'     from the running program, and its own message hooks - Desaware
'     SpyWorks' WinHook is a common one - end up sharing the thread's
'     hook chain with handlers that call back into the IDE.
'
'   - the tab-bar timer keeps firing inside the running program's
'     message loop, and every poll behind it (tab bar, line-number
'     gutter, bookmarks, git, navigation, auto-backup) either walks
'     the VBE object model or shells out. Re-entering the IDE while it
'     is executing project code is exactly what breaks it.
'
'  So everything automatic asks Mode_IsDesign() first.
'
'  The state comes from the IDE's own title bar, which always carries
'  it in the first brackets - "Project1 - Microsoft Visual Basic
'  [design]", "[run]", "[break]". VBProject.Mode looks like the proper
'  answer and is not: VB6 leaves it at vbext_vm_Run for a project
'  sitting in the designer, so gating on it switched the whole add-in
'  off. The title never lies, and reading it is a plain GetWindowText
'  - no COM call from inside a message hook.
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
' when no input is reaching the hook, and repaints the editor overlays
' once a run ends - nothing drew them while project code owned the
' thread, and no paint is coming on its own.
Public Sub Mode_Poll()
    On Error Resume Next
    Mode_Invalidate
    Dim d As Boolean
    d = Mode_IsDesign()
    If d = mWasDesign Then Exit Sub
    mWasDesign = d
    If d Then Highlight_InvalidateAll
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
