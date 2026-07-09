Attribute VB_Name = "modBackup"
Option Explicit

' =====================================================================
'  Auto-backup: zips the active project's folder into "<dir>\.backups"
'  every INTERVAL_MIN minutes, but only if some project file changed
'  since the last backup. The zip runs in a hidden PowerShell
'  (Compress-Archive, ships with Windows) so the IDE never waits.
'  Oldest archives are rotated out beyond KEEP_COUNT.
'  Toggle + "Backup Now" live in the Modernizr menu; the enabled flag
'  persists via SaveSetting.
' =====================================================================

Private Const INTERVAL_MIN As Long = 10
Private Const KEEP_COUNT As Long = 20
Private Const TICKS_PER_CHECK As Long = 75    ' ~30 s at 400 ms

Private mEnabled As Boolean
Private mTicks As Long
Private mLastBackup As Date

Public Sub Backup_Init()
    On Error Resume Next
    mEnabled = (GetSetting("VB6Modernizr", "Backup", "Enabled", "1") = "1")
    mLastBackup = CDate(Val(GetSetting("VB6Modernizr", "Backup", "Last", "0")))
End Sub

Public Function Backup_Enabled() As Boolean
    Backup_Enabled = mEnabled
End Function

Public Sub Backup_Toggle()
    On Error Resume Next
    mEnabled = Not mEnabled
    SaveSetting "VB6Modernizr", "Backup", "Enabled", IIf(mEnabled, "1", "0")
    MsgBox "Auto-backup is now " & IIf(mEnabled, "ON", "OFF") & "." & _
           vbCrLf & vbCrLf & "Every " & INTERVAL_MIN & " minutes (when " & _
           "files changed), keeping the last " & KEEP_COUNT & _
           " zips in the project's .backups folder.", _
           vbInformation, "VB6 Modernizr"
End Sub

' called from the tab-bar timer (~400 ms)
Public Sub Backup_Poll()
    On Error Resume Next
    mTicks = mTicks + 1
    If mTicks < TICKS_PER_CHECK Then Exit Sub
    mTicks = 0
    If Not mEnabled Then Exit Sub
    If DateDiff("n", mLastBackup, Now) < INTERVAL_MIN Then Exit Sub
    If Not ChangedSince(mLastBackup) Then Exit Sub
    Backup_Now False
End Sub

Public Sub Backup_Now(ByVal interactive As Boolean)
    On Error Resume Next
    Dim projDir As String
    projDir = ActiveProjDir()
    If Len(projDir) = 0 Then
        If interactive Then MsgBox "Save the project first.", _
            vbExclamation, "VB6 Modernizr"
        Exit Sub
    End If

    Dim bdir As String
    bdir = projDir & "\.backups"
    If Len(Dir$(bdir, vbDirectory)) = 0 Then MkDir bdir
    Rotate bdir

    Dim zip As String
    zip = bdir & "\backup_" & Format$(Now, "yyyymmdd_hhnnss") & ".zip"
    Shell "powershell -NoProfile -ExecutionPolicy Bypass -Command " & _
          """Get-ChildItem -LiteralPath '" & projDir & "' -Exclude " & _
          "'.backups' | Compress-Archive -DestinationPath '" & zip & _
          "' -Force""", vbHide

    mLastBackup = Now
    SaveSetting "VB6Modernizr", "Backup", "Last", CStr(CDbl(Now))
    If interactive Then
        MsgBox "Backup started (runs in the background):" & vbCrLf & zip, _
               vbInformation, "VB6 Modernizr"
    End If
End Sub

' ---------------------------------------------------------------------

Private Function ActiveProjDir() As String
    On Error Resume Next
    Dim f As String, p As Long
    f = gVBE.ActiveVBProject.FileName
    p = InStrRev(f, "\")
    If p > 0 Then ActiveProjDir = Left$(f, p - 1)
End Function

' any project file (or the .vbp itself) newer than t?
Private Function ChangedSince(ByVal t As Date) As Boolean
    On Error Resume Next
    Dim proj As VBIDE.VBProject, comp As VBIDE.VBComponent, j As Long
    Set proj = gVBE.ActiveVBProject
    If proj Is Nothing Then Exit Function
    If FileDateTime(proj.FileName) > t Then ChangedSince = True: Exit Function
    For Each comp In proj.VBComponents
        For j = 1 To comp.FileCount
            If FileDateTime(comp.FileNames(j)) > t Then
                ChangedSince = True
                Exit Function
            End If
        Next
    Next
End Function

' timestamped names sort chronologically; drop the oldest beyond quota
Private Sub Rotate(ByVal bdir As String)
    On Error Resume Next
    Dim f As String, names() As String, n As Long
    ReDim names(0 To 32)
    f = Dir$(bdir & "\backup_*.zip")
    Do While Len(f) > 0
        If n > UBound(names) Then ReDim Preserve names(0 To n * 2)
        names(n) = f
        n = n + 1
        f = Dir$()
    Loop
    If n < KEEP_COUNT Then Exit Sub

    Dim i As Long, j As Long, tmp As String
    For i = 0 To n - 2
        For j = i + 1 To n - 1
            If names(j) < names(i) Then
                tmp = names(i): names(i) = names(j): names(j) = tmp
            End If
        Next
    Next
    For i = 0 To n - KEEP_COUNT       ' leaves KEEP_COUNT-1 + the new one
        Kill bdir & "\" & names(i)
    Next
End Sub
