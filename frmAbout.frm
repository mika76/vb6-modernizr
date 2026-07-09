VERSION 5.00
Begin VB.Form frmAbout
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About VB6 Modernizr"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8100
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCaret
      Interval        =   530
      Left            =   7560
      Top             =   120
   End
   Begin VB.CommandButton cmdOK
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   420
      Left            =   6600
      TabIndex        =   0
      Top             =   2700
      Width           =   1300
   End
   Begin VB.Label lblVer
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   555
   End
   Begin VB.Label lblTag
      BackStyle       =   0  'Transparent
      Caption         =   "Modern text editing for the VB6 IDE. The missing ""e"" is a feature."
      Height          =   400
      Left            =   240
      TabIndex        =   2
      Top             =   2610
      Width           =   6200
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  About box: the logo lives in VB6Modernizr.res (bitmap 102, icon
'  group 101) so the DLL stays self-contained. The wordmark's text
'  cursor is real - tmrCaret blinks it over the bitmap.
' =====================================================================

' logo placement and the caret cell inside it, in pixels
Private Const LOGO_L As Long = 14
Private Const LOGO_T As Long = 14
Private Const CARET_X As Long = 451
Private Const CARET_Y As Long = 45
Private Const CARET_W As Long = 3
Private Const CARET_H As Long = 37

Private Const INK As Long = &H26231C    ' wordmark ink (BGR)

Private mCaretOn As Boolean

Public Sub ShowAbout()
    On Error Resume Next
    Load Me
    SetWindowLongA Me.hwnd, GWL_HWNDPARENT, MainHwnd()
    Me.Show
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Theme_ApplyIcon Me
    lblVer.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    Me.PaintPicture LoadResPicture(102, vbResBitmap), _
        LOGO_L * Screen.TwipsPerPixelX, LOGO_T * Screen.TwipsPerPixelY
End Sub

Private Sub tmrCaret_Timer()
    Dim x As Long, y As Long
    mCaretOn = Not mCaretOn
    x = (LOGO_L + CARET_X) * Screen.TwipsPerPixelX
    y = (LOGO_T + CARET_Y) * Screen.TwipsPerPixelY
    Me.Line (x, y)-Step(CARET_W * Screen.TwipsPerPixelX, CARET_H * Screen.TwipsPerPixelY), _
        IIf(mCaretOn, INK, vbWhite), BF
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub
