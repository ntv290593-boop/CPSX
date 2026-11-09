Attribute VB_Name = "modSrcLinks"
'SRC_LINKS add-in: bam link trong file co sheet _LINKS -> mo dung trang PDF (Edge) / dung o xlsx.
'O co nhieu link -> hien menu chon.  _LINKS: A Key(Sheet!O) B Sheet C O D Nhan E File(tuong doi) F Trang G O dich
'Tuy chon: H1 = "BASE_DIR", H2 = thu muc file luc build. Chuyen file di cho khac van tim lai duoc (xem SL_ResolvePath).
Option Explicit

Private mApp As clsAppEvents
Private mWb As Workbook
Private Const LINKS_SHEET As String = "_LINKS"
Private Const POPUP_NAME As String = "SRC_LINKS_POPUP"
Private Const BASE_HDR As String = "BASE_DIR"

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
    full = SL_ResolvePath(wb, r)
    pg = Trim$(CStr(ws.Cells(r, 6).Value))
    subA = Trim$(CStr(ws.Cells(r, 7).Value))
    If Not fso.FileExists(full) Then
        'May chua dong bo: mo thang tren web, khong kiem tra truoc duoc file co ton tai khong
        If LCase$(Left$(wb.Path, 4)) = "http" Then
            full = SL_WebUrl(wb.Path, CStr(ws.Cells(r, 5).Value))
            Select Case LCase$(fso.GetExtensionName(full))
            Case "xlsx", "xlsm", "xls", "xlsb"
                SL_GotoWb full, subA
            Case Else
                full = SL_UrlEncode(full)
                If Len(pg) > 0 Then full = full & "#page=" & pg
                CreateObject("Shell.Application").ShellExecute "msedge.exe", """" & full & """", "", "open", 1
            End Select
            Exit Sub
        End If
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

'Tim file theo thu tu: (1) tuong doi theo cho dat workbook; (2) tuong doi theo BASE_DIR luc build;
'(3) bo cac "..\" roi thu tung thu muc cha cua workbook (workbook chuyen trong cay du an, hoac ca du an doi o/doi may).
'File mo tu SharePoint/OneDrive co Path la https://: doi sang thu muc dong bo tren may truoc khi tim.
'Khong thay thi tra ve duong dan (1) de bao loi.
Public Function SL_ResolvePath(ByVal wb As Workbook, ByVal r As Long) As String
    Dim ws As Worksheet, fso As Object, rel As String, tail As String, d As String, first As String, cand As String, goc As String
    Set ws = wb.Worksheets(LINKS_SHEET)
    Set fso = CreateObject("Scripting.FileSystemObject")
    rel = Replace(Trim$(CStr(ws.Cells(r, 5).Value)), "/", "\")
    goc = wb.Path
    If LCase$(Left$(goc, 4)) = "http" Then
        d = SL_UrlToLocal(goc)
        If Len(d) > 0 Then goc = d
    End If
    first = SL_Join(fso, goc, rel)
    If SL_Exists(fso, first) Then SL_ResolvePath = first: Exit Function
    If CStr(ws.Range("H1").Value) = BASE_HDR Then
        cand = SL_Join(fso, CStr(ws.Range("H2").Value), rel)
        If SL_Exists(fso, cand) Then SL_ResolvePath = cand: Exit Function
    End If
    tail = rel
    Do While Left$(tail, 3) = "..\"
        tail = Mid$(tail, 4)
    Loop
    'Van con https:// = may nay chua dong bo, khong do theo thu muc cha duoc
    If Len(tail) > 0 And LCase$(Left$(goc, 4)) <> "http" Then
        d = goc
        Do While Len(d) > 0
            cand = fso.BuildPath(d, tail)
            If SL_Exists(fso, cand) Then SL_ResolvePath = cand: Exit Function
            d = fso.GetParentFolderName(d)
        Loop
    End If
    SL_ResolvePath = first
End Function

'Doi dia chi https cua thu muc sang thu muc dong bo tren may, doc tu danh sach thu vien OneDrive da dong bo.
'(1) UrlNamespace dai nhat la dau cua dia chi. (2) Thu muc nguoi khac chia se, them bang "Add shortcut to OneDrive":
'dia chi khong khop, nen thu cac doan cuoi cua dia chi ngay duoi moi thu muc goc. Khong thay thi tra ve "".
Public Function SL_UrlToLocal(ByVal url As String) As String
    Const HKCU As Long = &H80000001
    Const KEY_OD As String = "Software\SyncEngines\Providers\OneDrive"
    Dim reg As Object, fso As Object, names As Variant, n As Variant, ns As Variant, mp As Variant
    Dim u As String, ns2 As String, best As String, bestMp As String, cand As String, mounts As New Collection
    Dim segs() As String, i As Long, j As Long, m As Variant, duoi As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    u = Replace(url, "%20", " ")
    If Right$(u, 1) <> "/" Then u = u & "/"
    'Du phong khi WMI/registry bi chan: thu muc OneDrive lay tu bien moi truong (doc Unicode qua WScript.Shell)
    On Error Resume Next
    Dim sh As Object, ev As Variant, evp As String
    Set sh = CreateObject("WScript.Shell")
    For Each ev In Array("OneDriveCommercial", "OneDrive", "OneDriveConsumer")
        evp = sh.ExpandEnvironmentStrings("%" & ev & "%")
        If Len(evp) > 0 And Left$(evp, 1) <> "%" Then
            If fso.FolderExists(evp) Then mounts.Add evp
        End If
    Next ev
    Set reg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    If Not reg Is Nothing Then reg.EnumKey HKCU, KEY_OD, names
    On Error GoTo Ket
    If Not IsArray(names) Then names = Array()
    For Each n In names
        ns = Null: mp = Null
        reg.GetStringValue HKCU, KEY_OD & "\" & n, "UrlNamespace", ns
        reg.GetStringValue HKCU, KEY_OD & "\" & n, "MountPoint", mp
        If Not IsNull(mp) Then
            If Len(mp) > 0 Then
                If Not IsNull(ns) Then
                    If Len(ns) > 0 Then
                        mounts.Add CStr(mp)
                        ns2 = CStr(ns)
                        If Right$(ns2, 1) <> "/" Then ns2 = ns2 & "/"
                        If LCase$(Left$(u, Len(ns2))) = LCase$(ns2) And Len(ns2) > Len(best) Then best = ns2: bestMp = CStr(mp)
                    End If
                End If
            End If
        End If
    Next n
    If Len(best) > 0 Then
        cand = fso.BuildPath(bestMp, Replace(Mid$(u, Len(best) + 1), "/", "\"))
        If Right$(cand, 1) = "\" Then cand = Left$(cand, Len(cand) - 1)
        If fso.FolderExists(cand) Then SL_UrlToLocal = cand: Exit Function
    End If
    'segs: bo https:, "", ten mien va "/" cuoi
    segs = Split(Left$(u, Len(u) - 1), "/")
    For i = 3 To UBound(segs)
        duoi = ""
        For j = i To UBound(segs)
            duoi = duoi & "\" & segs(j)
        Next j
        For Each m In mounts
            If fso.FolderExists(m & duoi) Then SL_UrlToLocal = m & duoi: Exit Function
        Next m
    Next i
Ket:
End Function

'Dia chi web cua file dich khi may chua dong bo. Doan dau cua duong dan co trong dia chi workbook thi noi tai do
'(workbook da chuyen vao trong cay du an); khong thi di len dung so "..\" nhu duong dan tuong doi.
Public Function SL_WebUrl(ByVal wbPath As String, ByVal rel As String) As String
    Dim u As String, tail As String, segs() As String, dau As String, i As Long, n As Long
    u = Replace(wbPath, "%20", " ")
    If Right$(u, 1) = "/" Then u = Left$(u, Len(u) - 1)
    tail = Replace(Trim$(rel), "/", "\")
    Do While Left$(tail, 3) = "..\"
        tail = Mid$(tail, 4): n = n + 1
    Loop
    segs = Split(u, "/")
    dau = Split(tail, "\")(0)
    For i = UBound(segs) To 3 Step -1
        If StrComp(segs(i), dau, vbTextCompare) = 0 Then
            ReDim Preserve segs(i - 1)
            SL_WebUrl = Join(segs, "/") & "/" & Replace(tail, "\", "/")
            Exit Function
        End If
    Next i
    If UBound(segs) - n >= 3 Then
        ReDim Preserve segs(UBound(segs) - n)
        u = Join(segs, "/")
    End If
    SL_WebUrl = u & "/" & Replace(tail, "\", "/")
End Function

Private Function SL_Join(ByVal fso As Object, ByVal d As String, ByVal rel As String) As String
    On Error Resume Next
    SL_Join = fso.GetAbsolutePathName(fso.BuildPath(d, rel))
    If Err.Number <> 0 Then SL_Join = d & "\" & rel
End Function

Private Function SL_Exists(ByVal fso As Object, ByVal p As String) As Boolean
    On Error Resume Next
    SL_Exists = fso.FileExists(p)
End Function

Private Sub SL_GotoWb(ByVal full As String, ByVal subA As String)
    Dim w As Workbook, twb As Workbook, p As Long, shName As String, cellA As String, ten As String
    'So theo ten: file mo tu web co FullName https, ma Excel khong mo duoc 2 file trung ten
    ten = Replace(full, "\", "/")
    ten = Mid$(ten, InStrRev(ten, "/") + 1)
    For Each w In Application.Workbooks
        If StrComp(w.Name, ten, vbTextCompare) = 0 Then Set twb = w: Exit For
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
