VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents XL As Application

Private Sub XL_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Hyperlink)
    On Error GoTo EH
    modSrcLinks.SL_HandleClick Sh, Target
    Exit Sub
EH:
    MsgBox "SRC_LINKS: " & Err.Description, vbExclamation
End Sub
