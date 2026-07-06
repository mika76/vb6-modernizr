VERSION 5.00
Begin VB.Form frmRefs
   Caption         =   "References"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstRefs
      Height          =   3400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8160
   End
   Begin VB.Label lblStatus
      Caption         =   ""
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   8160
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

Private Const LB_SETHORIZONTALEXTENT As Long = &H194

Private mRefs() As MatchInfo
Private mCount As Long

Public Sub ShowRefs(ByVal word As String)
    On Error Resume Next
    Load Me
    SendMessageA lstRefs.hwnd, LB_SETHORIZONTALEXTENT, 3000, 0
    SetWindowLongA Me.hwnd, GWL_HWNDPARENT, MainHwnd()

    mRefs = gMatches
    mCount = gMatchCount

    Me.Caption = "References to '" & word & "'"
    lstRefs.Clear
    Dim i As Long
    For i = 0 To mCount - 1
        lstRefs.AddItem mRefs(i).Comp & "(" & mRefs(i).LineNum & "): " & _
                        Left$(Trim$(mRefs(i).LineText), 250)
    Next
    lblStatus.Caption = mCount & " reference(s) in the active project."

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

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    lstRefs.Width = Me.ScaleWidth - 240
    lstRefs.Height = Me.ScaleHeight - lstRefs.Top - 500
    lblStatus.Top = Me.ScaleHeight - 380
    lblStatus.Width = Me.ScaleWidth - 240
End Sub
