Attribute VB_Name = "modActions"
Option Explicit

' =====================================================================
'  Central command dispatcher: menu items (clsMenuButton) and any
'  future toolbar buttons route here by action name.
'
'  Toggle menu items (tabs / guides / backup) register their
'  CommandBarButton here so their checked state can track the flag.
' =====================================================================

Private Const msoButtonUp As Long = 0
Private Const msoButtonDown As Long = -1

Private mToggleBtns As Collection   ' CommandBarButtons keyed by action

Public Sub Menu_RegisterToggle(ByVal act As String, btn As Object)
    On Error Resume Next
    If mToggleBtns Is Nothing Then Set mToggleBtns = New Collection
    mToggleBtns.Add btn, act
End Sub

Public Sub Menu_ClearToggles()
    Set mToggleBtns = Nothing
End Sub

' Reflect the current on/off flags as menu check marks.
Public Sub Menu_SyncToggles()
    On Error Resume Next
    SetToggleState "tabs", gTabBarVisible
    SetToggleState "guides", gGuidesEnabled
    SetToggleState "linenums", gLineNumsEnabled
    SetToggleState "backup", Backup_Enabled()
End Sub

Private Sub SetToggleState(ByVal act As String, ByVal onOff As Boolean)
    On Error Resume Next
    If mToggleBtns Is Nothing Then Exit Sub
    Dim btn As Object
    Set btn = mToggleBtns(act)
    If btn Is Nothing Then Exit Sub
    btn.State = IIf(onOff, msoButtonDown, msoButtonUp)
End Sub

Public Sub DoAction(ByVal act As String)
    On Error Resume Next
    Select Case act
    Case "findbar":   frmFind.ShowBar
    Case "findfiles": frmFindFiles.ShowDialog
    Case "refs":      Edit_FindAllReferences
    Case "def":       Edit_GoToDefinition
    Case "hlword":    Edit_HighlightWord
    Case "browser":   frmBrowser.ShowBrowser
    Case "bmtoggle":  BM_Toggle
    Case "bmnext":    BM_NextBookmark
    Case "bmclear":   BM_ClearAll
    Case "gitchanges": frmChanges.ShowChanges
    Case "gitlog":    frmGitLog.ShowLog
    Case "gitblame":  Git_BlameCurrentLine
    Case "dup":       Edit_DuplicateLines
    Case "moveup":    Edit_MoveLinesUp
    Case "movedown":  Edit_MoveLinesDown
    Case "delline":   Edit_DeleteLines
    Case "comment":   Edit_ToggleComment
    Case "clearhl":   Highlight_Clear
    Case "tabs":      TabBar_Toggle
    Case "guides":    Guides_Toggle
    Case "linenums":  LineNums_Toggle
    Case "backup":    Backup_Toggle
    Case "backupnow": Backup_Now True
    Case "keys":      frmShortcuts.ShowSheet
    Case "about":     frmAbout.ShowAbout
    End Select

    Menu_SyncToggles
End Sub
