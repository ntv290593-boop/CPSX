Attribute VB_Name = "modCpsxLib"
Option Explicit
' CPSX_LIB: cai thu vien JS (cpsx_lib.js) vao sheet an _LIB cua workbook costing
' de Claude trong Excel nap 1 dong thay vi viet lai script moi lan.
Private Const LIB_PATH As String = "D:\COSTING EXCEL\_MASTER\cpsx-master\lib\cpsx_lib.js"
Private Const CHUNK As Long = 30000

' Chay tay: Alt+F8 > PERSONAL.XLSB!InstallCpsxLib (cai/cap nhat cho file dang mo)
Public Sub InstallCpsxLib()
    If ActiveWorkbook Is Nothing Then Exit Sub
    Dim v As String: v = WriteLib(ActiveWorkbook, True)
    If v <> "" Then MsgBox "Da cai CPSX_LIB v" & v & " vao sheet _LIB (an)." & vbLf & "Nho Ctrl+S de luu.", vbInformation, "CPSX_LIB"
End Sub

' Goi tu clsAppEvents khi mo file: chi cap nhat khi file la workbook costing va _LIB cu/thieu
Public Sub AutoUpdateLib(ByVal wb As Workbook)
    On Error GoTo Quit
    If Not HasSheet(wb, "Chiet tinh") Then Exit Sub
    Dim txt As String: txt = ReadUtf8(LIB_PATH)
    If txt = "" Then Exit Sub
    If HasSheet(wb, "_LIB") Then
        If CStr(wb.Worksheets("_LIB").Range("B1").Value) = LibVersion(txt) Then Exit Sub
    End If
    WriteLib wb, False
Quit:
End Sub

Private Function WriteLib(ByVal wb As Workbook, ByVal loud As Boolean) As String
    Dim txt As String: txt = ReadUtf8(LIB_PATH)
    If txt = "" Then
        If loud Then MsgBox "Khong doc duoc " & LIB_PATH, vbCritical, "CPSX_LIB"
        Exit Function
    End If
    Dim ws As Worksheet, cur As Worksheet, scr As Boolean
    scr = Application.ScreenUpdating: Application.ScreenUpdating = False
    Set cur = ActiveSheet
    If HasSheet(wb, "_LIB") Then
        Set ws = wb.Worksheets("_LIB")
        ws.Visible = xlSheetVisible
        ws.Cells.Clear
    Else
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "_LIB"
    End If
    ws.Columns("A").NumberFormat = "@"
    Dim pos As Long, cut As Long, r As Long, n As Long
    n = Len(txt): pos = 1: r = 1
    Do While pos <= n
        cut = pos + CHUNK - 1
        If cut >= n Then
            cut = n
        Else
            cut = InStrRev(txt, vbLf, cut)
            If cut < pos Then cut = pos + CHUNK - 1
        End If
        ws.Cells(r, 1).Value = Mid$(txt, pos, cut - pos + 1)
        pos = cut + 1: r = r + 1
    Loop
    ws.Range("B1").NumberFormat = "@"
    ws.Range("B1").Value = LibVersion(txt)
    ws.Range("B2").Value = "Thu vien JS cho Claude trong Excel - khong sua tay. Cap nhat: PERSONAL.XLSB!InstallCpsxLib"
    ws.Visible = xlSheetVeryHidden
    On Error Resume Next
    cur.Activate
    Application.ScreenUpdating = scr
    WriteLib = LibVersion(txt)
End Function

Private Function LibVersion(ByVal txt As String) As String
    Dim p As Long, q As Long
    p = InStr(1, txt, "VER: """)
    If p = 0 Then LibVersion = "?": Exit Function
    p = p + 6: q = InStr(p, txt, """")
    LibVersion = Mid$(txt, p, q - p)
End Function

Private Function HasSheet(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim s As Object
    On Error Resume Next
    Set s = wb.Worksheets(nm)
    HasSheet = Not s Is Nothing
End Function

Private Function ReadUtf8(ByVal path As String) As String
    On Error GoTo Fail
    If Dir(path) = "" Then Exit Function
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.LoadFromFile path
    ReadUtf8 = st.ReadText(-1)
    st.Close
    Exit Function
Fail:
    ReadUtf8 = ""
End Function
