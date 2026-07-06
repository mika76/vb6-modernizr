Attribute VB_Name = "modMRU"
Option Explicit

' =====================================================================
'  Most-recently-used order of MDI windows, for the Ctrl+Tab switcher.
'  Keys are NormalizeCaption(caption) & "|" & window type, the same
'  identity the tab bar uses. frmTabs feeds MRU_Touch whenever the
'  active window changes.
' =====================================================================

Private mKeys As Collection

Public gSwitcherActive As Boolean

Public Sub MRU_Touch(ByVal key As String)
    On Error Resume Next
    If Len(key) = 0 Then Exit Sub
    If mKeys Is Nothing Then Set mKeys = New Collection
    mKeys.Remove key
    Err.Clear
    If mKeys.Count = 0 Then
        mKeys.Add key, key
    Else
        mKeys.Add key, key, 1      ' most recent first
    End If
End Sub

Public Function MRU_Count() As Long
    If Not mKeys Is Nothing Then MRU_Count = mKeys.Count
End Function

Public Function MRU_Key(ByVal i As Long) As String
    On Error Resume Next
    MRU_Key = mKeys(i)
End Function
