VERSION 5.00
Begin VB.Form frmRefs
   Caption         =   "References"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB6Modernizr.ucList lstRefs
      Height          =   3300
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   5821
   End
   Begin VB.Label lblStatus
      Caption         =   ""
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   3780
      Width           =   8040
   End
End
Attribute VB_Name = "frmRefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
'  "Find All References" results (Shift+F12). Snapshot of the matches
'  from modSearch; double-click jumps to the reference.
' =====================================================================

Private mRefs() As MatchInfo
Private mCount As Long

Public Sub ShowRefs(ByVal word As String, _
        Optional ByVal verb As String = "References to")
    On Error Resume Next
    Load Me
    SetWindowLongA Me.hwnd, GWL_HWNDPARENT, MainHwnd()

    mRefs = gMatches
    mCount = gMatchCount

    Me.Caption = verb & " '" & word & "'"
    lstRefs.Clear
    Dim i As Long
    For i = 0 To mCount - 1
        lstRefs.AddItem mRefs(i).Comp & "(" & mRefs(i).LineNum & "): " & _
                        Left$(Trim$(mRefs(i).LineText), 250), , _
                        FileForComponent(mRefs(i).Comp)
    Next
    lblStatus.Caption = mCount & " item(s) in the active project."

    Me.Show vbModeless
    If lstRefs.ListCount > 0 Then lstRefs.ListIndex = 0
End Sub

Private Sub lstRefs_DblClick()
    On Error Resume Next
    Dim i As Long
    i = lstRefs.ListIndex
    If i < 0 Or i >= mCount Then Exit Sub
    GoToMatch mRefs(i)
End Sub

Private Sub lstRefs_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        lstRefs_DblClick
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        KeyCode = 0
        Me.Hide
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    lstRefs.Width = Me.ScaleWidth - MARGIN_STD * 2
    lstRefs.Height = Me.ScaleHeight - lstRefs.Top - 560
    lblStatus.Top = Me.ScaleHeight - 420
    lblStatus.Width = Me.ScaleWidth - MARGIN_STD * 2
End Sub

Private Sub Form_Load()
    Theme_ApplyIcon Me
End Sub
