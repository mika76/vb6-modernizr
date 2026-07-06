VERSION 5.00
Begin VB.Form frmShortcuts
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modernizr - Keyboard Shortcuts"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstKeys
      BeginProperty Font
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4740
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6560
   End
   Begin VB.CommandButton cmdClose
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   360
      Left            =   5560
      TabIndex        =   1
      Top             =   5100
      Width           =   1120
   End
End
Attribute VB_Name = "frmShortcuts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Cheat sheet for everything the add-in binds. Ctrl+Shift+/ in a code
' window, or Modernizr -> Keyboard Shortcuts.

Public Sub ShowSheet()
    On Error Resume Next
    Load Me
    SetWindowLongA Me.hwnd, GWL_HWNDPARENT, MainHwnd()
    Me.Show vbModeless
End Sub

Private Sub Form_Load()
    Populate
End Sub

Private Sub Populate()
    lstKeys.Clear
    Hdr "FIND"
    Row "Ctrl+F", "Open find / replace bar"
    Row "Enter / Shift+Enter", "Find next / previous (in the bar)"
    Row "F3 / Shift+F3", "Find next / previous (anywhere)"
    Row "Ctrl+F3", "Highlight word at cursor everywhere"
    Row "Esc", "Close find bar"
    Hdr ""
    Hdr "NAVIGATE  (focus in code window)"
    Row "F12", "Go to definition of word at cursor"
    Row "Shift+F12", "Find all references to word at cursor"
    Row "Ctrl+F2", "Toggle bookmark on current line"
    Row "F2", "Next bookmark"
    Row "Ctrl+Shift+O", "Code browser (procedures / TODOs)"
    Hdr ""
    Hdr "EDIT  (focus in code window)"
    Row "Ctrl+D", "Duplicate line / selection"
    Row "Alt+Up / Alt+Down", "Move lines up / down"
    Row "Ctrl+Shift+K", "Delete lines"
    Row "Ctrl+/", "Comment / uncomment"
    Hdr ""
    Hdr "WINDOWS"
    Row "Ctrl+Tab", "Switch window (hold Ctrl, Tab cycles,"
    Row "", "  release Ctrl to switch, Esc cancels)"
    Row "Ctrl+Shift+Tab", "Switch window, backwards"
    Row "Middle-click tab", "Close window"
    Row "Right-click tab", "Close / Close Others / Close All"
    Hdr ""
    Hdr "GIT  (repo auto-detected from the project folder)"
    Row "Ctrl+Shift+G", "Changes / commit window"
    Row "Ctrl+Shift+B", "Blame current line"
    Row "", "Tab bar shows branch; margin bars show"
    Row "", "  changed lines (green=add, blue=edit)"
    Hdr ""
    Hdr "MOUSE"
    Row "Wheel", "Scroll code window"
    Row "Shift+Wheel", "Scroll horizontally"
    Hdr ""
    Hdr "HELP"
    Row "Ctrl+Shift+/", "This window"
End Sub

Private Sub Hdr(ByVal s As String)
    lstKeys.AddItem s
End Sub

Private Sub Row(ByVal keys As String, ByVal desc As String)
    lstKeys.AddItem "  " & Left$(keys & Space$(22), 22) & desc
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
