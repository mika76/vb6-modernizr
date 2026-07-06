Attribute VB_Name = "modActions"
Option Explicit

' =====================================================================
'  Central command dispatcher: menu items (clsMenuButton) and any
'  future toolbar buttons route here by action name.
' =====================================================================

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
    Case "keys":      frmShortcuts.ShowSheet
    Case "about":
        MsgBox "VB6 Modernizr 1.1" & vbCrLf & vbCrLf & _
               "MDI window tabs, find/replace bar with highlighting," & vbCrLf & _
               "find in files, references, editing shortcuts and" & vbCrLf & _
               "mouse wheel scrolling for the VB6 IDE.", _
               vbInformation, "VB6 Modernizr"
    End Select
End Sub
