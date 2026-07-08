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
Private mButtons As Collection        ' of clsMenuButton (event sinks)
Private mConnected As Boolean

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
    Wheel_Init
    Guides_Init
    Backup_Init
    TabBar_Init
    Menu_SyncToggles
End Sub

Private Sub TermAddin()
    On Error Resume Next
    Highlight_Terminate
    Wheel_Term
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

    Theme_FreeIcons

    If Not mMenuPopup Is Nothing Then mMenuPopup.Delete
    Set mMenuPopup = Nothing
    Set mButtons = Nothing
    Menu_ClearToggles

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
    AddBtn "Auto-Backu&p", "backup", False
    AddBtn "Backup No&w", "backupnow", False

    AddBtn "&Keyboard Shortcuts... (Ctrl+Shift+/)", "keys", True
    AddBtn "&About VB6 Modernizr", "about", False
End Sub

Private Sub AddBtn(ByVal cap As String, ByVal act As String, _
        ByVal group As Boolean)
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
    Case "tabs", "guides", "backup"
        Menu_RegisterToggle act, btn
    End Select
End Sub
