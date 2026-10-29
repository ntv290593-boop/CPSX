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
' Bat su kien Save cua MOI workbook dang mo -> day quy tac CHUNG len Master
Public WithEvents App As Application

Private Sub App_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    On Error Resume Next
    modRulesSync.PushSharedRules Wb, True
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    On Error Resume Next
    modCpsxLib.AutoUpdateLib Wb
End Sub
