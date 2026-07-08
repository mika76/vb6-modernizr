Attribute VB_Name = "modTheme"
Option Explicit

' =====================================================================
'  Shared visual theme: colors, spacing, and the shell icon cache.
'
'  Icons come from SHGetFileInfo with SHGFI_USEFILEATTRIBUTES, so a
'  component that has never been saved still gets the icon registered
'  for its extension. HICONs are cached per extension and destroyed
'  once, at disconnect (Theme_FreeIcons).
' =====================================================================

' --- palette (BGR) ----------------------------------------------------

Public Const THEME_CODE As Long = &HB45A00      ' blue = code module
Public Const THEME_DESIGN As Long = &H3C8C1E    ' green = designer/form
Public Const THEME_ACCENT As Long = &H157DE9    ' orange = modified/match
Public Const THEME_BORDER As Long = &H808080    ' popup frame gray

' --- spacing -----------------------------------------------------------

Public Const MARGIN_STD As Long = 180           ' twips, dialog edge margin

' --- shell icon cache state --------------------------------------------

Private mIcons As Collection    ' HICON keyed by "i" & extension

' --- git graph lanes ----------------------------------------------------

Public Function LaneColor(ByVal lane As Long) As Long
    Select Case ((lane Mod 8) + 8) Mod 8
        Case 0: LaneColor = &HB45A00    ' blue
        Case 1: LaneColor = &H3C8C1E    ' green
        Case 2: LaneColor = &H157DE9    ' orange
        Case 3: LaneColor = &HB4008C    ' purple
        Case 4: LaneColor = &H2020C8    ' red
        Case 5: LaneColor = &H968C00    ' teal
        Case 6: LaneColor = &H8C14DC    ' pink
        Case Else: LaneColor = &H1E78A0 ' brown
    End Select
End Function

' --- shell icon cache ---------------------------------------------------

' Small (16px) shell icon for a file path or bare name; cached by
' extension. Works for files that do not exist on disk.
Public Function IconForFile(ByVal sPath As String) As Long
    On Error Resume Next
    Dim ext As String, p As Long
    p = InStrRev(sPath, ".")
    If p > 0 And p > InStrRev(sPath, "\") Then ext = LCase$(Mid$(sPath, p))

    If mIcons Is Nothing Then Set mIcons = New Collection
    Err.Clear
    IconForFile = mIcons("i" & ext)
    If Err.Number = 0 Then Exit Function
    Err.Clear

    Dim sfi As SHFILEINFO
    SHGetFileInfoA "x" & ext, FILE_ATTRIBUTE_NORMAL, sfi, Len(sfi), _
        SHGFI_ICON Or SHGFI_SMALLICON Or SHGFI_USEFILEATTRIBUTES
    mIcons.Add sfi.hIcon, "i" & ext
    IconForFile = sfi.hIcon
End Function

' File name (or "x.<ext>" placeholder) for a VBComponent, resolved by
' name across open projects. Prefers the saved file's real extension,
' falls back to the component type. Feed the result to IconForFile.
Public Function FileForComponent(ByVal compName As String) As String
    On Error Resume Next
    Dim c As VBIDE.VBComponent, prj As VBIDE.VBProject, f As String

    For Each prj In gVBE.VBProjects
        Err.Clear
        Set c = prj.VBComponents(compName)
        If Err.Number = 0 And Not c Is Nothing Then Exit For
        Set c = Nothing
    Next

    f = ""
    If Not c Is Nothing Then
        Err.Clear
        f = c.FileNames(1)
        If Err.Number <> 0 Then f = ""
        If Len(f) = 0 Then f = "x" & ExtForCompType(c.Type)
    Else
        f = "x.bas"
    End If
    FileForComponent = f
End Function

Public Function IconForComponent(ByVal compName As String) As Long
    IconForComponent = IconForFile(FileForComponent(compName))
End Function

' Icon for an IDE window caption like "frmTabs (Code)".
Public Function IconForCaption(ByVal cap As String) As Long
    Dim p As Long
    p = InStrRev(cap, " (")
    If p > 0 Then cap = Left$(cap, p - 1)
    IconForCaption = IconForComponent(Trim$(cap))
End Function

Private Function ExtForCompType(ByVal t As Long) As String
    Select Case t
        Case vbext_ct_ClassModule: ExtForCompType = ".cls"
        Case vbext_ct_VBForm, vbext_ct_VBMDIForm, vbext_ct_MSForm
            ExtForCompType = ".frm"
        Case vbext_ct_UserControl: ExtForCompType = ".ctl"
        Case vbext_ct_PropPage: ExtForCompType = ".pag"
        Case vbext_ct_ActiveXDesigner: ExtForCompType = ".dsr"
        Case Else: ExtForCompType = ".bas"
    End Select
End Function

' Draw a cached icon at 16px (DPI scaled) on any hDC.
Public Sub DrawIcon16(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
        ByVal hIcon As Long)
    If hIcon = 0 Then Exit Sub
    Dim s As Long
    s = ScaleForDpi(16)
    DrawIconEx hdc, x, y, hIcon, s, s, 0, 0, DI_NORMAL
End Sub

Public Sub Theme_FreeIcons()
    On Error Resume Next
    Dim v As Variant
    If mIcons Is Nothing Then Exit Sub
    For Each v In mIcons
        If CLng(v) <> 0 Then DestroyIcon CLng(v)
    Next
    Set mIcons = Nothing
End Sub
