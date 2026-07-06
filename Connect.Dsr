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
'  Wires the add-in into the IDE: menu items, tab bar, wheel hook.
' =====================================================================

Public VBInstance As VBIDE.VBE

Private mMenuPopup As Object          ' Office CommandBarPopup (late bound)
Private mBtnFind As Object
Private mBtnFindFiles As Object
Private mBtnTabs As Object
Private mBtnClearHL As Object
Private mBtnAbout As Object

Private WithEvents evtFind As VBIDE.CommandBarEvents
Attribute evtFind.VB_VarHelpID = -1
Private WithEvents evtFindFiles As VBIDE.CommandBarEvents
Attribute evtFindFiles.VB_VarHelpID = -1
Private WithEvents evtTabs As VBIDE.CommandBarEvents
Attribute evtTabs.VB_VarHelpID = -1
Private WithEvents evtClearHL As VBIDE.CommandBarEvents
Attribute evtClearHL.VB_VarHelpID = -1
Private WithEvents evtAbout As VBIDE.CommandBarEvents
Attribute evtAbout.VB_VarHelpID = -1

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
    TabBar_Init
End Sub

Private Sub TermAddin()
    On Error Resume Next
    Highlight_Terminate
    Wheel_Term
    TabBar_Term
    Unhook_All

    Unload frmFind
    Unload frmFindFiles

    If Not mMenuPopup Is Nothing Then mMenuPopup.Delete
    Set mMenuPopup = Nothing
    Set evtFind = Nothing: Set evtFindFiles = Nothing
    Set evtTabs = Nothing: Set evtClearHL = Nothing: Set evtAbout = Nothing

    Set gVBE = Nothing
    Set VBInstance = Nothing
    mConnected = False
End Sub

Private Sub AddMenus()
    On Error Resume Next
    Dim cbMenuBar As Object
    Set cbMenuBar = VBInstance.CommandBars("Menu Bar")
    If cbMenuBar Is Nothing Then Exit Sub

    ' 10 = msoControlPopup, 1 = msoControlButton; Temporary:=True
    Set mMenuPopup = cbMenuBar.Controls.Add(10, , , , True)
    mMenuPopup.Caption = "M&odernizr"

    Set mBtnFind = mMenuPopup.Controls.Add(1)
    mBtnFind.Caption = "&Find / Replace Bar (Ctrl+F)"
    Set evtFind = VBInstance.Events.CommandBarEvents(mBtnFind)

    Set mBtnFindFiles = mMenuPopup.Controls.Add(1)
    mBtnFindFiles.Caption = "Find in Fi&les..."
    Set evtFindFiles = VBInstance.Events.CommandBarEvents(mBtnFindFiles)

    Set mBtnClearHL = mMenuPopup.Controls.Add(1)
    mBtnClearHL.Caption = "&Clear Highlights"
    Set evtClearHL = VBInstance.Events.CommandBarEvents(mBtnClearHL)

    Set mBtnTabs = mMenuPopup.Controls.Add(1)
    mBtnTabs.Caption = "Show/Hide &Tabs"
    mBtnTabs.BeginGroup = True
    Set evtTabs = VBInstance.Events.CommandBarEvents(mBtnTabs)

    Set mBtnAbout = mMenuPopup.Controls.Add(1)
    mBtnAbout.Caption = "&About VB6 Modernizr"
    mBtnAbout.BeginGroup = True
    Set evtAbout = VBInstance.Events.CommandBarEvents(mBtnAbout)
End Sub

' ---------------------------------------------------------------------

Private Sub evtFind_Click(ByVal CommandBarControl As Object, _
        handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    frmFind.ShowBar
    handled = True
End Sub

Private Sub evtFindFiles_Click(ByVal CommandBarControl As Object, _
        handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    frmFindFiles.ShowDialog
    handled = True
End Sub

Private Sub evtTabs_Click(ByVal CommandBarControl As Object, _
        handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    TabBar_Toggle
    handled = True
End Sub

Private Sub evtClearHL_Click(ByVal CommandBarControl As Object, _
        handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    Highlight_Clear
    handled = True
End Sub

Private Sub evtAbout_Click(ByVal CommandBarControl As Object, _
        handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    MsgBox "VB6 Modernizr 1.0" & vbCrLf & vbCrLf & _
           "MDI window tabs, Find/Replace with match highlighting," & vbCrLf & _
           "Find in Files, and mouse wheel scrolling for the VB6 IDE.", _
           vbInformation, "VB6 Modernizr"
    handled = True
End Sub
