Attribute VB_Name = "modSrcLinks"
'SRC_LINKS add-in: bam link trong file co sheet _LINKS -> mo dung trang PDF (Edge) / dung o xlsx.
'O co nhieu link -> hien menu chon.  _LINKS: A Key(Sheet!O) B Sheet C O D Nhan E File(tuong doi) F Trang G O dich
Option Explicit

Private mApp As clsAppEvents
Private mWb As Workbook
Private Const LINKS_SHEET As String = "_LINKS"
Private Const POPUP_NAME As String = "SRC_LINKS_POPUP"

Public Sub SL_Init()
    Set mApp = New clsAppEvents
    Set mApp.XL = Application
End Sub

Public Sub SL_HandleClick(ByVal Sh As Object, ByVal Target As Hyperlink)
    Dim wb As Workbook, ws As Worksheet, key As String, rows As Collection
    Set wb = Sh.Parent
    On Error Resume Next
    Set ws = wb.Worksheets(LINKS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    If Sh.Name = LINKS_SHEET Then Exit Sub
    key = Sh.Name & "!" & Target.Range.Cells(1, 1).Address(False, False)
    Set rows = SL_FindRows(ws, key)
    If rows.Count = 0 Then Exit Sub
    If rows.Count = 1 Then
        SL_OpenRow wb, CLng(rows(1))
    Else
        SL_ShowMenu wb, ws, rows
    End If
End Sub

Private Function SL_FindRows(ByVal ws As Worksheet, ByVal key As String) As Collection
    Dim res As New Collection, m As Variant, r As Long, lastR As Long
    m = Application.Match(key, ws.Columns(1), 0)
    If Not IsError(m) Then
        lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        r = CLng(m)
        Do While r <= lastR
            If CStr(ws.Cells(r, 1).Value) <> key Then Exit Do
            res.Add r
            r = r + 1
        Loop
    End If
    Set SL_FindRows = res
End Function

Private Sub SL_ShowMenu(ByVal wb As Workbook, ByVal ws As Worksheet, ByVal rows As Collection)
    Dim cb As Object, b As Object, r As Variant
    On Error Resume Next
    Application.CommandBars(POPUP_NAME).Delete
    On Error GoTo 0
    Set cb = Application.CommandBars.Add(Name:=POPUP_NAME, Position:=5, Temporary:=True)
    Set mWb = wb
    For Each r In rows
        Set b = cb.Controls.Add(Type:=1)
        b.Caption = CStr(ws.Cells(r, 4).Value)
        b.Parameter = CStr(r)
        b.OnAction = "'" & ThisWorkbook.Name & "'!SL_MenuClick"
    Next r
    cb.ShowPopup
End Sub

Public Sub SL_MenuClick()
    If mWb Is Nothing Then Exit Sub
    SL_OpenRow mWb, CLng(Application.CommandBars.ActionControl.Parameter)
End Sub

Public Sub SL_OpenRow(ByVal wb As Workbook, ByVal r As Long)
    Dim ws As Worksheet, fso As Object, full As String, pg As String, subA As String, ext As String
    Set ws = wb.Worksheets(LINKS_SHEET)
    Set fso = CreateObject("Scripting.FileSystemObject")
    full = fso.GetAbsolutePathName(fso.BuildPath(wb.Path, CStr(ws.Cells(r, 5).Value)))
    pg = Trim$(CStr(ws.Cells(r, 6).Value))
    subA = Trim$(CStr(ws.Cells(r, 7).Value))
    If Not fso.FileExists(full) Then
        MsgBox "Khong tim thay file:" & vbLf & full, vbExclamation, "SRC_LINKS"
        Exit Sub
    End If
    ext = LCase$(fso.GetExtensionName(full))
    Select Case ext
    Case "pdf"
        Dim url As String
        url = "file:///" & SL_UrlEncode(Replace(full, "\", "/"))
        If Len(pg) > 0 Then url = url & "#page=" & pg
        CreateObject("Shell.Application").ShellExecute "msedge.exe", """" & url & """", "", "open", 1
    Case "xlsx", "xlsm", "xls", "xlsb"
        SL_GotoWb full, subA
    Case Else
        CreateObject("Shell.Application").ShellExecute full, "", "", "open", 1
    End Select
End Sub

Private Sub SL_GotoWb(ByVal full As String, ByVal subA As String)
    Dim w As Workbook, twb As Workbook, p As Long, shName As String, cellA As String
    For Each w In Application.Workbooks
        If StrComp(w.FullName, full, vbTextCompare) = 0 Then Set twb = w: Exit For
    Next w
    If twb Is Nothing Then Set twb = Application.Workbooks.Open(Filename:=full, UpdateLinks:=0, ReadOnly:=True)
    twb.Activate
    If Len(subA) = 0 Then Exit Sub
    p = InStrRev(subA, "!")
    shName = Replace(Left$(subA, p - 1), "'", "")
    cellA = Mid$(subA, p + 1)
    On Error Resume Next
    Application.Goto twb.Worksheets(shName).Range(cellA), True
    If Err.Number <> 0 Then MsgBox "Khong toi duoc o: " & subA, vbExclamation, "SRC_LINKS"
End Sub

Private Function SL_UrlEncode(ByVal s As String) As String
    Dim i As Long, j As Long, c As Long, ch As String, out As String, b() As Byte
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        c = AscW(ch) And &HFFFF&
        If (c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or InStr("-_.~/:", ch) > 0 Then
            out = out & ch
        Else
            b = SL_Utf8(c)
            For j = 0 To UBound(b)
                out = out & "%" & Right$("0" & Hex$(b(j)), 2)
            Next j
        End If
    Next i
    SL_UrlEncode = out
End Function

Private Function SL_Utf8(ByVal c As Long) As Byte()
    Dim r() As Byte
    If c < &H80& Then
        ReDim r(0): r(0) = c
    ElseIf c < &H800& Then
        ReDim r(1): r(0) = &HC0 Or (c \ &H40&): r(1) = &H80 Or (c And &H3F&)
    Else
        ReDim r(2): r(0) = &HE0 Or (c \ &H1000&): r(1) = &H80 Or ((c \ &H40&) And &H3F&): r(2) = &H80 Or (c And &H3F&)
    End If
    SL_Utf8 = r
End Function
