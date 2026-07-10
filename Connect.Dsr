VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   6000
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   6585
   _ExtentX        =   11615
   _ExtentY        =   10583
   _Version        =   393216
   Description     =   "Modern IDE helpers for VB6: MDI tabs, find/replace with highlighting, find in files, mouse wheel scrolling."
   DisplayName     =   "VB6 Modernizr"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' =====================================================================
'  VB6 Modernizr - add-in entry point
'  Wires the add-in into the IDE: menu, tab bar, message hook.
'  Menu clicks are dispatched by action name via clsMenuButton ->
'  modActions.DoAction, so adding a command is a single AddBtn line.
' =====================================================================

Public VBInstance As VBIDE.VBE

Private mMenuPopup As Object          ' Office CommandBarPopup (late bound)
Private mToolBar As Object            ' Office CommandBar (late bound)
Private mButtons As Collection        ' of clsMenuButton (event sinks)
Private mConnected As Boolean

' Office CommandBar constants (late bound, so not in a type library)
Private Const msoBarTop As Long = 1
Private Const msoBarFloating As Long = 4
Private Const msoButtonIcon As Long = 1
Private Const msoButtonCaption As Long = 2

' ---------------------------------------------------------------------

Private Sub AddinInstance_OnConnection(ByVal Application As Object, _
        ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, _
        ByVal AddInInst As Object, custom() As Variant)
    On Error Resume Next
    Set VBInstance = Application
    Set gVBE = VBInstance
    ' At IDE startup the MDI client may not be fully created yet;
    ' wait for OnStartupComplete in that case.
    If ConnectMode <> ext_cm_Startup Then InitAddin
End Sub

Private Sub AddinInstance_OnStartupComplete(custom() As Variant)
    On Error Resume Next
    InitAddin
End Sub

Private Sub AddinInstance_OnDisconnection( _
        ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, _
        custom() As Variant)
    On Error Resume Next
    TermAddin
End Sub

' ---------------------------------------------------------------------

Private Sub InitAddin()
    On Error Resume Next
    If mConnected Then Exit Sub
    mConnected = True

    AddMenus
    AddToolbar
    Wheel_Init
    Guides_Init
    LineNums_Init
    Backup_Init
    TabBar_Init
    Menu_SyncToggles
End Sub

Private Sub TermAddin()
    On Error Resume Next
    Highlight_Terminate
    Wheel_Term
    ' release the gutter strip before the MDI hook goes away
    gLineNumsEnabled = False
    TabBar_Term
    Unhook_All

    Unload frmFind
    Unload frmFindFiles
    Unload frmRefs
    Unload frmSwitcher
    Unload frmShortcuts
    Unload frmBrowser
    Unload frmChanges
    Unload frmGitLog
    Unload frmGutter
    Unload frmAbout

    Theme_FreeIcons

    If Not mMenuPopup Is Nothing Then mMenuPopup.Delete
    Set mMenuPopup = Nothing
    If Not mToolBar Is Nothing Then
        SaveToolbarPos
        mToolBar.Delete
    End If
    Set mToolBar = Nothing
    Set mButtons = Nothing
    Menu_ClearToggles
    Nav_Term

    Set gVBE = Nothing
    Set VBInstance = Nothing
    mConnected = False
End Sub

' ---------------------------------------------------------------------

Private Sub AddMenus()
    On Error Resume Next
    Dim cbMenuBar As Object
    Set cbMenuBar = VBInstance.CommandBars("Menu Bar")
    If cbMenuBar Is Nothing Then Exit Sub

    Set mButtons = New Collection

    ' 10 = msoControlPopup; Temporary:=True
    Set mMenuPopup = cbMenuBar.Controls.Add(10, , , , True)
    mMenuPopup.Caption = "M&odernizr"

    AddBtn "&Find / Replace Bar (Ctrl+F)", "findbar", False
    AddBtn "Find in Fi&les...", "findfiles", False
    AddBtn "Find All &References (Shift+F12)", "refs", False
    AddBtn "&Go to Definition (F12)", "def", False
    AddBtn "Highlight Word at Cursor (Ctrl+F3)", "hlword", False
    AddBtn "Code Bro&wser... (Ctrl+Shift+O)", "browser", False
    AddBtn "Go to F&ile... (Ctrl+P)", "quickopen", False

    Dim mnuBack As Object, mnuFwd As Object
    Set mnuBack = AddBtn("&Back (Alt+Left)", "navback", True)
    Set mnuFwd = AddBtn("For&ward (Alt+Right)", "navfwd", False)
    Nav_RegisterButtons mnuBack, mnuFwd

    AddBtn "Toggle Book&mark (Ctrl+F2)", "bmtoggle", True
    AddBtn "Next Bookmark (F2)", "bmnext", False
    AddBtn "Clear All Bookmarks", "bmclear", False

    AddBtn "&Duplicate Line (Ctrl+D)", "dup", True
    AddBtn "Move Lines &Up (Alt+Up)", "moveup", False
    AddBtn "Move Lines Do&wn (Alt+Down)", "movedown", False
    AddBtn "Delete Li&nes (Ctrl+Shift+K)", "delline", False
    AddBtn "&Comment / Uncomment (Ctrl+/)", "comment", False

    AddBtn "Git Chan&ges / Commit... (Ctrl+Shift+G)", "gitchanges", True
    AddBtn "Git Log... (Ctrl+Shift+L)", "gitlog", False
    AddBtn "Git Blame Line (Ctrl+Shift+B)", "gitblame", False

    AddBtn "Clear &Highlights", "clearhl", True
    AddBtn "Show/Hide &Tabs", "tabs", False
    AddBtn "Indentation G&uides", "guides", False
    AddBtn "Line &Numbers", "linenums", False
    AddBtn "Auto-Backu&p", "backup", False
    AddBtn "Backup No&w", "backupnow", False

    AddBtn "&Keyboard Shortcuts... (Ctrl+Shift+/)", "keys", True
    AddBtn "&About VB6 Modernizr", "about", False
End Sub

Private Function AddBtn(ByVal cap As String, ByVal act As String, _
        ByVal group As Boolean) As Object
    On Error Resume Next
    Dim btn As Object, mb As clsMenuButton
    Set btn = mMenuPopup.Controls.Add(1)      ' msoControlButton
    btn.Caption = cap
    If group Then btn.BeginGroup = True
    Set mb = New clsMenuButton
    mb.Action = act
    Set mb.evt = VBInstance.Events.CommandBarEvents(btn)
    mButtons.Add mb

    ' on/off commands show a check mark tracking their flag
    Select Case act
    Case "tabs", "guides", "linenums", "backup"
        Menu_RegisterToggle act, btn
    End Select
    Set AddBtn = btn
End Function

' ---------------------------------------------------------------------
'  Toolbar: the everyday commands plus the on/off toggles (which show
'  as pressed buttons). Created temporary so nothing is orphaned if
'  the add-in goes away - which is also why the IDE won't remember
'  where the user docked it; position is saved to the registry at
'  disconnect and re-applied here instead.
' ---------------------------------------------------------------------

Private Sub AddToolbar()
    On Error Resume Next
    ' a leftover bar from a crashed session would make Add fail
    Set mToolBar = VBInstance.CommandBars("Modernizr")
    If Not mToolBar Is Nothing Then mToolBar.Delete
    Set mToolBar = Nothing

    Set mToolBar = VBInstance.CommandBars.Add("Modernizr", msoBarTop, , True)
    If mToolBar Is Nothing Then Exit Sub

    ' icon buttons use Office's built-in FaceId glyphs; 0 = text button
    Dim btnBack As Object, btnFwd As Object
    Set btnBack = AddTool("Back (Alt+Left)", "navback", 1017, False)
    Set btnFwd = AddTool("Forward (Alt+Right)", "navfwd", 1018, False)
    Nav_RegisterButtons btnBack, btnFwd

    AddTool "Go to File (Ctrl+P)", "quickopen", 23, True
    AddTool "Find / Replace Bar (Ctrl+F)", "findbar", 141, False
    AddTool "Find in Files", "findfiles", 172, False
    AddTool "Code Browser (Ctrl+Shift+O)", "browser", 225, False

    AddTool "Git Changes / Commit (Ctrl+Shift+G)", "gitchanges", 0, True, "Git"
    AddTool "Git Log (Ctrl+Shift+L)", "gitlog", 0, False, "Log"

    AddTool "Show/Hide Tabs", "tabs", 0, True, "Tabs"
    AddTool "Indentation Guides", "guides", 0, False, "Guides"
    AddTool "Line Numbers", "linenums", 226, False
    AddTool "Auto-Backup", "backup", 3, False

    RestoreToolbarPos
    mToolBar.Visible = _
        (GetSetting("VB6Modernizr", "Toolbar", "Visible", "1") = "1")
End Sub

Private Function AddTool(ByVal tip As String, ByVal act As String, _
        ByVal faceId As Long, ByVal group As Boolean, _
        Optional ByVal cap As String = "") As Object
    On Error Resume Next
    Dim btn As Object, mb As clsMenuButton
    Set btn = mToolBar.Controls.Add(1)        ' msoControlButton
    If faceId > 0 Then
        btn.faceId = faceId
        btn.style = msoButtonIcon
    Else
        btn.style = msoButtonCaption
    End If
    btn.Caption = IIf(Len(cap) > 0, cap, tip)
    btn.ToolTipText = tip
    If group Then btn.BeginGroup = True
    Set mb = New clsMenuButton
    mb.Action = act
    Set mb.evt = VBInstance.Events.CommandBarEvents(btn)
    mButtons.Add mb

    Select Case act
    Case "tabs", "guides", "linenums", "backup"
        Menu_RegisterToggle act, btn
    End Select
    Set AddTool = btn
End Function

Private Sub RestoreToolbarPos()
    On Error Resume Next
    Dim p As Long
    p = CLng(GetSetting("VB6Modernizr", "Toolbar", "Position", _
                        CStr(msoBarTop)))
    mToolBar.Position = p
    If p = msoBarFloating Then
        mToolBar.Top = CLng(GetSetting("VB6Modernizr", "Toolbar", "Top", "200"))
        mToolBar.Left = CLng(GetSetting("VB6Modernizr", "Toolbar", "Left", "200"))
    Else
        Dim r As Long
        r = CLng(GetSetting("VB6Modernizr", "Toolbar", "Row", "0"))
        If r > 0 Then mToolBar.RowIndex = r
        mToolBar.Left = CLng(GetSetting("VB6Modernizr", "Toolbar", "Left", "0"))
    End If
End Sub

Private Sub SaveToolbarPos()
    On Error Resume Next
    SaveSetting "VB6Modernizr", "Toolbar", "Position", CStr(mToolBar.Position)
    SaveSetting "VB6Modernizr", "Toolbar", "Row", CStr(mToolBar.RowIndex)
    SaveSetting "VB6Modernizr", "Toolbar", "Left", CStr(mToolBar.Left)
    SaveSetting "VB6Modernizr", "Toolbar", "Top", CStr(mToolBar.Top)
    SaveSetting "VB6Modernizr", "Toolbar", "Visible", _
                IIf(mToolBar.Visible, "1", "0")
End Sub
