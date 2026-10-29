Attribute VB_Name = "modAiInput"
Option Explicit
'===============================================================
' AI_INPUT: Cowork ghi <du an>\02_takeoff\AI_INPUT.xlsx (bang tblAiInput)
' -> Power Query qAiInput nap vao sheet AI_INPUT (bang tblAiInputView)
' File CPSX nam o <du an>\03_costing\  => duong dan tu tinh theo file.
' Dat trong PERSONAL.XLSB. Goi SyncAiInput tu clsAppEvents.App_WorkbookOpen
'===============================================================
Public Const AI_SHEET As String = "AI_INPUT"
Public Const AI_VIEW As String = "tblAiInputView"
Public Const AI_SRC_TABLE As String = "tblAiInput"
Public Const AI_QUERY As String = "qAiInput"
Public Const AI_PARAM As String = "pAiInputPath"
Public Const AI_COLS As String = "Ma,DongBoQ,Ten,Loai,W,D,H,SL,VatLieu,HoanThien,ThongSo,Nguon,DoTinCay,Trang,GhiChu"

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

Private Function MainFormula() As String
    Dim cols As String, c As Variant
    For Each c In Split(AI_COLS, ",")
        cols = cols & IIf(cols = "", "", ", ") & """" & c & """"
    Next c
    MainFormula = "let" & vbCrLf & _
        "    Cols = {" & cols & "}," & vbCrLf & _
        "    Empty = #table(Cols, {})," & vbCrLf & _
        "    Src = try Excel.Workbook(File.Contents(" & AI_PARAM & "), true) otherwise null," & vbCrLf & _
        "    Tbl = if Src = null then Empty else try Src{[Item=""" & AI_SRC_TABLE & """, Kind=""Table""]}[Data] otherwise Empty," & vbCrLf & _
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

' Chay 1 lan tren template (Alt+F8 > PERSONAL.XLSB!SetupAiInput)
Public Sub SetupAiInput()
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, p As String
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    p = AiPath(wb)
    If p = "" Then p = "D:\COSTING EXCEL\_ACTIVE\02_takeoff\AI_INPUT.xlsx"
    UpsertQ wb, AI_PARAM, ParamFormula(p)
    UpsertQ wb, AI_QUERY, MainFormula()

    On Error Resume Next
    Set ws = wb.Worksheets(AI_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = AI_SHEET
        ws.Tab.Color = RGB(255, 192, 0)
    End If
    Set lo = FindLo(wb, AI_VIEW)
    If lo Is Nothing Then
        ws.Cells.Clear
        Set lo = ws.ListObjects.Add(SourceType:=0, _
            Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & AI_QUERY & ";Extended Properties=""""", _
            Destination:=ws.Range("A1"))
        With lo.QueryTable
            .CommandType = xlCmdSql
            .CommandText = Array("SELECT * FROM [" & AI_QUERY & "]")
            .RefreshOnFileOpen = False
            .BackgroundQuery = False
            .Refresh BackgroundQuery:=False
        End With
        lo.Name = AI_VIEW
    Else
        lo.QueryTable.Refresh BackgroundQuery:=False
    End If
    MsgBox "OK: " & AI_SHEET & " / " & AI_QUERY & vbLf & "Nguon: " & p, vbInformation, "AI_INPUT"
End Sub

' Goi khi mo file: cap nhat duong dan theo vi tri file + lam moi neu file ton tai
Public Sub SyncAiInput(ByVal wb As Workbook)
    On Error GoTo Quit
    If Not HasQuery(wb, AI_QUERY) Then Exit Sub
    Dim p As String: p = AiPath(wb)
    If p = "" Then Exit Sub
    UpsertQ wb, AI_PARAM, ParamFormula(p)
    If Dir(p) <> "" Then
        Dim lo As ListObject: Set lo = FindLo(wb, AI_VIEW)
        If Not lo Is Nothing Then lo.QueryTable.Refresh BackgroundQuery:=False
    End If
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
    Dim lo As ListObject: Set lo = FindLo(wb, AI_VIEW)
    lo.QueryTable.Refresh BackgroundQuery:=False
    Dim n As Long: If Not lo.DataBodyRange Is Nothing Then n = lo.ListRows.Count
    MsgBox "Da nap " & n & " dong tu" & vbLf & p, vbInformation, "AI_INPUT"
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
