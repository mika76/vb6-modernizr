VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Event sink for one menu button: forwards clicks to DoAction, so the
' Connect designer doesn't need a WithEvents variable per command.

Public Action As String
Public WithEvents evt As VBIDE.CommandBarEvents
Attribute evt.VB_VarHelpID = -1

Private Sub evt_Click(ByVal CommandBarControl As Object, _
        handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    DoAction Action
    handled = True
End Sub
