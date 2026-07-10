Attribute VB_Name = "modActions"
Option Explicit

' =====================================================================
'  Central command dispatcher: menu items (clsMenuButton) and any
'  future toolbar buttons route here by action name.
'
'  Toggle commands (tabs / guides / backup) register their
'  CommandBarButtons here so their checked state can track the flag;
'  one action can have several buttons (menu item + toolbar button).
' =====================================================================

Private Const msoButtonUp As Long = 0
Private Const msoButtonDown As Long = -1

Private mToggleActs As Collection   ' action names, parallel to...
Private mToggleBtns As Collection   ' ...their CommandBarButtons

Public Sub Menu_RegisterToggle(ByVal act As String, btn As Object)
    On Error Resume Next
    If mToggleBtns Is Nothing Then
        Set mToggleBtns = New Collection
        Set mToggleActs = New Collection
    End If
    mToggleBtns.Add btn
    mToggleActs.Add act
End Sub

Public Sub Menu_ClearToggles()
    Set mToggleBtns = Nothing
    Set mToggleActs = Nothing
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
    Dim i As Long
    For i = 1 To mToggleActs.Count
        If mToggleActs(i) = act Then
            mToggleBtns(i).State = IIf(onOff, msoButtonDown, msoButtonUp)
        End If
    Next
End Sub

Public Sub DoAction(ByVal act As String)
    On Error Resume Next
    Select Case act

    ' window commands toggle: a second click (menu or toolbar) while
    ' the window is showing closes it again
    Case "findbar"
        If gFindBarVisible Then frmFind.HideBar Else frmFind.ShowBar
    Case "findfiles"
        If FormShowing("frmFindFiles") Then
            frmFindFiles.Hide
        Else
            frmFindFiles.ShowDialog
        End If
    Case "browser"
        If FormShowing("frmBrowser") Then
            frmBrowser.Hide
        Else
            frmBrowser.ShowBrowser
        End If
    Case "quickopen"
        If FormShowing("frmQuickOpen") Then
            frmQuickOpen.Hide
        ElseIf Not frmQuickOpen.JustDismissed Then
            ' clicking its toolbar button deactivates (= hides) the
            ' palette before this runs - don't instantly reopen it
            frmQuickOpen.ShowQuickOpen
        End If
    Case "gitchanges"
        If FormShowing("frmChanges") Then
            frmChanges.Hide
        Else
            frmChanges.ShowChanges
        End If
    Case "gitlog"
        If FormShowing("frmGitLog") Then
            frmGitLog.Hide
        Else
            frmGitLog.ShowLog
        End If
    Case "keys"
        If FormShowing("frmShortcuts") Then
            frmShortcuts.Hide
        Else
            frmShortcuts.ShowSheet
        End If

    Case "refs":      Edit_FindAllReferences
    Case "def":       Edit_GoToDefinition
    Case "hlword":    Edit_HighlightWord
    Case "navback":   Nav_Back
    Case "navfwd":    Nav_Forward
    Case "bmtoggle":  BM_Toggle
    Case "bmnext":    BM_NextBookmark
    Case "bmclear":   BM_ClearAll
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
    Case "about":     frmAbout.ShowAbout
    End Select

    Menu_SyncToggles
End Sub

' Loaded AND visible - iterating Forms avoids the side effect of
' auto-loading a predeclared form just to ask about it.
Private Function FormShowing(ByVal nm As String) As Boolean
    On Error Resume Next
    Dim f As Object
    For Each f In Forms
        If StrComp(f.Name, nm, vbTextCompare) = 0 Then
            FormShowing = f.Visible
            Exit Function
        End If
    Next
End Function
