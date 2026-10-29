Attribute VB_Name = "modAiInput"
Option Explicit
'===============================================================
' AI_INPUT: Cowork ghi <du an>\02_takeoff\AI_INPUT.xlsx (bang tblAiInput)
' -> Power Query qAiInput nap vao sheet AI_INPUT (bang tblAiInputView)
' v2 (2026-10-24): khoa = BlockId; them bang tblAiModule (1 dong = 1 o nhap)
'    -> qAiModule -> sheet AI_MODULE (bang tblAiModuleView); Claude goi L.fromAiInput
' File CPSX nam o <du an>\03_costing\  => duong dan tu tinh theo file.
' Dat trong PERSONAL.XLSB. Goi SyncAiInput tu clsAppEvents.App_WorkbookOpen
'===============================================================
Public Const AI_SHEET As String = "AI_INPUT"
Public Const AI_VIEW As String = "tblAiInputView"
Public Const AI_SRC_TABLE As String = "tblAiInput"
Public Const AI_QUERY As String = "qAiInput"
Public Const AI_PARAM As String = "pAiInputPath"
Public Const AI_COLS As String = "BlockId,ThaoTac,Ma,DongBoQ,QuanHe,Ten,Loai,Mau,MauTen,W,D,H,SL,GiaTheo,VatLieu,HoanThien,ThongSo,Nguon,DoTinCay,Trang,GhiChu,RunId"
Public Const MOD_SHEET As String = "AI_MODULE"
Public Const MOD_VIEW As String = "tblAiModuleView"
Public Const MOD_SRC_TABLE As String = "tblAiModule"
Public Const MOD_QUERY As String = "qAiModule"
Public Const MOD_COLS As String = "BlockId,Module,Lan,Off,Cot,GiaTri,Loai,Nguon,Trang,DoTinCay,GhiChu"

' Duong dan AI_INPUT.xlsx cho workbook: <thu muc cha cua 03_costing>\02_takeoff\AI_INPUT.xlsx
Public Function AiPath(ByVal wb As Workbook) As String
    Dim p As String: p = wb.Path
    If p = "" Then Exit Function
    If InStrRev(p, "\") > 0 Then p = Left$(p, InStrRev(p, "\") - 1)
    AiPath = p & "\02_takeoff\AI_INPUT.xlsx"
End Function

Private Function ParamFormula(ByVal path As String) As String
    ParamFormula = """" & path & """ meta [IsParameterQuery=true, Type=""Text"", IsParameterQueryRequired=true]"
End Function

Private Function MainFormula(Optional ByVal colList As String = AI_COLS, Optional ByVal srcTable As String = AI_SRC_TABLE) As String
    Dim cols As String, c As Variant
    For Each c In Split(colList, ",")
        cols = cols & IIf(cols = "", "", ", ") & """" & c & """"
    Next c
    MainFormula = "let" & vbCrLf & _
        "    Cols = {" & cols & "}," & vbCrLf & _
        "    Empty = #table(Cols, {})," & vbCrLf & _
        "    Src = try Excel.Workbook(File.Contents(" & AI_PARAM & "), true) otherwise null," & vbCrLf & _
        "    Tbl = if Src = null then Empty else try Src{[Item=""" & srcTable & """, Kind=""Table""]}[Data] otherwise Empty," & vbCrLf & _
        "    Out = Table.SelectRows(Tbl, each List.NonNullCount(Record.FieldValues(_)) > 0)" & vbCrLf & _
        "in" & vbCrLf & _
        "    Out"
End Function

Private Sub UpsertQ(wb As Workbook, ByVal qName As String, ByVal formula As String)
    Dim q As Object
    On Error Resume Next
    Set q = wb.Queries(qName)
    On Error GoTo 0
    If q Is Nothing Then
        wb.Queries.Add Name:=qName, Formula:=formula
    ElseIf q.Formula <> formula Then
        q.Formula = formula
    End If
End Sub

Private Function FindLo(wb As Workbook, ByVal nm As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = nm Then Set FindLo = lo: Exit Function
        Next lo
    Next ws
End Function

Private Function HasQuery(wb As Workbook, ByVal qName As String) As Boolean
    Dim q As Object
    On Error Resume Next
    Set q = wb.Queries(qName)
    HasQuery = Not q Is Nothing
End Function

' Tao sheet + bang noi query neu chua co; da co thi lam moi
Private Sub EnsureView(wb As Workbook, ByVal sheetName As String, ByVal viewName As String, ByVal qName As String)
    Dim ws As Worksheet, lo As ListObject
    Set lo = FindLo(wb, viewName)
    If Not lo Is Nothing Then
        lo.QueryTable.Refresh BackgroundQuery:=False
        Exit Sub
    End If
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
        ws.Tab.Color = RGB(255, 192, 0)
    End If
    ws.Cells.Clear
    Set lo = ws.ListObjects.Add(SourceType:=0, _
        Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & qName & ";Extended Properties=""""", _
        Destination:=ws.Range("A1"))
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & qName & "]")
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .Refresh BackgroundQuery:=False
    End With
    lo.Name = viewName
End Sub

Private Function RefreshViews(wb As Workbook) As Long
    Dim v As Variant, lo As ListObject
    For Each v In Array(AI_VIEW, MOD_VIEW)
        Set lo = FindLo(wb, CStr(v))
        If Not lo Is Nothing Then
            lo.QueryTable.Refresh BackgroundQuery:=False
            If v = AI_VIEW And Not lo.DataBodyRange Is Nothing Then RefreshViews = lo.ListRows.Count
        End If
    Next v
End Function

' Chay 1 lan tren template (Alt+F8 > PERSONAL.XLSB!SetupAiInput). File v1: chay lai de them AI_MODULE
Public Sub SetupAiInput()
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    MsgBox "OK: " & AI_SHEET & " + " & MOD_SHEET & vbLf & "Nguon: " & SetupAiInputWb(wb), vbInformation, "AI_INPUT"
End Sub

' Ban khong hop thoai (goi tu script COM: Application.Run "PERSONAL.XLSB!SetupAiInputWb", wb)
Public Function SetupAiInputWb(ByVal wb As Workbook) As String
    Dim p As String: p = AiPath(wb)
    If p = "" Then p = "D:\COSTING EXCEL\_ACTIVE\02_takeoff\AI_INPUT.xlsx"
    UpsertQ wb, AI_PARAM, ParamFormula(p)
    UpsertQ wb, AI_QUERY, MainFormula()
    UpsertQ wb, MOD_QUERY, MainFormula(MOD_COLS, MOD_SRC_TABLE)
    EnsureView wb, AI_SHEET, AI_VIEW, AI_QUERY
    EnsureView wb, MOD_SHEET, MOD_VIEW, MOD_QUERY
    SetupAiInputWb = p
End Function

' Goi khi mo file: cap nhat duong dan theo vi tri file + lam moi neu file ton tai
Public Sub SyncAiInput(ByVal wb As Workbook)
    On Error GoTo Quit
    If Not HasQuery(wb, AI_QUERY) Then Exit Sub
    Dim p As String: p = AiPath(wb)
    If p = "" Then Exit Sub
    UpsertQ wb, AI_PARAM, ParamFormula(p)
    If Dir(p) <> "" Then RefreshViews wb
Quit:
End Sub

' Chay tay khi Cowork vua ghi xong AI_INPUT.xlsx (Alt+F8 > PERSONAL.XLSB!RefreshAiInput)
Public Sub RefreshAiInput()
    Dim wb As Workbook: Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If Not HasQuery(wb, AI_QUERY) Then
        MsgBox "File chua co AI_INPUT. Chay PERSONAL.XLSB!SetupAiInput", vbExclamation, "AI_INPUT"
        Exit Sub
    End If
    Dim p As String: p = AiPath(wb)
    UpsertQ wb, AI_PARAM, ParamFormula(p)
    If Dir(p) = "" Then
        MsgBox "Khong thay file: " & p, vbExclamation, "AI_INPUT"
        Exit Sub
    End If
    If FindLo(wb, MOD_VIEW) Is Nothing Then
        MsgBox "File chua co AI_MODULE (v2). Chay PERSONAL.XLSB!SetupAiInput", vbExclamation, "AI_INPUT"
        Exit Sub
    End If
    Dim n As Long: n = RefreshViews(wb)
    MsgBox "Da nap " & n & " block tu" & vbLf & p, vbInformation, "AI_INPUT"
End Sub

' Cai AI_INPUT vao template chuan (1 lan): Alt+F8 > PERSONAL.XLSB!SetupTemplateAiInput
Public Sub SetupTemplateAiInput()
    Const TEMPLATE_PATH As String = "D:\COSTING EXCEL\260903_DA_CPSX_RE00.xlsm"
    Dim wb As Workbook
    Set wb = Workbooks.Open(TEMPLATE_PATH)
    wb.Activate
    SetupAiInput
    wb.Save
End Sub
