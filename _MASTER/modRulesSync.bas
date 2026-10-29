Attribute VB_Name = "modRulesSync"
Option Explicit
'===============================================================
' Dong bo quy tac CHUNG giua file du an (tblRules) va RULES_MASTER
' Dat trong PERSONAL.XLSB. Chu Viet viet dang \uXXXX -> UX() giai ma
' (VBE khong luu duoc Unicode trong ma nguon).
'===============================================================

Public Const MASTER_PATH As String = "D:\COSTING EXCEL\_MASTER\RULES_MASTER.xlsx"
Public Const MASTER_TABLE As String = "tblRulesMaster"
Public Const PROJ_SHEET As String = "_RULES"
Public Const PROJ_TABLE As String = "tblRules"
Public Const VIEW_SHEET As String = "_RULES_MASTER"
Public Const VIEW_TABLE As String = "tblRulesMasterView"
Public Const SCOPE_SHARED As String = "CHUNG"
Public Const SCOPE_LOCAL As String = "RIENG"
Public Const STATUS_NEW As String = "Active"

Private gApp As clsAppEvents
Private gBusy As Boolean

'---------- Khoi dong (goi tu ThisWorkbook.Workbook_Open cua PERSONAL) ----------
Public Sub StartRulesSync()
    Set gApp = New clsAppEvents
    Set gApp.App = Application
End Sub

'---------- Ten cot ----------
Private Function H_ID() As String: H_ID = "ID": End Function
Private Function H_DATE() As String: H_DATE = UX("Ng\u00E0y"): End Function
Private Function H_TOPIC() As String: H_TOPIC = UX("Ch\u1EE7 \u0111\u1EC1"): End Function
Private Function H_RULE() As String: H_RULE = UX("Quy t\u1EAFc"): End Function
Private Function H_ORIGIN() As String: H_ORIGIN = UX("D\u1EF1 \u00E1n g\u1ED1c"): End Function
Private Function H_STATUS() As String: H_STATUS = UX("Tr\u1EA1ng th\u00E1i"): End Function
Private Function H_VER() As String: H_VER = UX("Phi\u00EAn b\u1EA3n"): End Function
Private Function H_SCOPE() As String: H_SCOPE = UX("Ph\u1EA1m vi"): End Function

Public Function UX(ByVal s As String) As String
    Dim p As Long
    p = InStr(1, s, "\u")
    Do While p > 0
        s = Left$(s, p - 1) & ChrW(CLng("&H" & Mid$(s, p + 2, 4))) & Mid$(s, p + 6)
        p = InStr(p + 1, s, "\u")
    Loop
    UX = s
End Function

'---------- Tien ich bang ----------
Private Function FindTable(wb As Workbook, ByVal tblName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tblName, vbTextCompare) = 0 Then Set FindTable = lo: Exit Function
        Next lo
    Next ws
End Function

Private Function ColIdx(lo As ListObject, ByVal header As String) As Long
    Dim c As ListColumn
    For Each c In lo.ListColumns
        If StrComp(Trim$(c.Name), header, vbTextCompare) = 0 Then ColIdx = c.Index: Exit Function
    Next c
End Function

' Tra ve dong trong de ghi; dung lai dong trong duy nhat cua bang moi tao
Private Function NewRow(lo As ListObject) As Range
    If lo.ListRows.Count = 1 Then
        If Application.WorksheetFunction.CountA(lo.DataBodyRange) = 0 Then
            Set NewRow = lo.DataBodyRange.Rows(1): Exit Function
        End If
    End If
    Set NewRow = lo.ListRows.Add.Range
End Function

Private Function EnsureColumn(lo As ListObject, ByVal header As String) As Long
    EnsureColumn = ColIdx(lo, header)
    If EnsureColumn = 0 Then
        With lo.ListColumns.Add
            .Name = header
            EnsureColumn = .Index
        End With
    End If
End Function

'---------- Buoc 1: tao file Master neu chua co ----------
Public Sub CreateMasterFile()
    Dim folder As String, wb As Workbook, ws As Worksheet, lo As ListObject, h As Variant
    If Dir(MASTER_PATH) <> "" Then Exit Sub
    folder = Left$(MASTER_PATH, InStrRev(MASTER_PATH, "\") - 1)
    If Dir(folder, vbDirectory) = "" Then MkDir folder
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "RULES"
    h = Array(H_ID, H_DATE, H_TOPIC, H_RULE, H_ORIGIN, H_STATUS, H_VER)
    ws.Range("A1").Resize(1, 7).Value = h
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:G2"), , xlYes)
    lo.Name = MASTER_TABLE
    wb.SaveAs MASTER_PATH, FileFormat:=51
    wb.Close False
End Sub

'---------- Buoc 2+3: chay 1 lan tren file du an dang mo ----------
Public Sub SetupProjectRules()
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, cScope As Long, cRule As Long, cell As Range
    Set wb = ActiveWorkbook
    Set lo = FindTable(wb, PROJ_TABLE)
    If lo Is Nothing Then
        On Error Resume Next
        Set ws = wb.Worksheets(PROJ_SHEET)
        On Error GoTo 0
        If ws Is Nothing Then
            Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
            ws.Name = PROJ_SHEET
        End If
        ws.Range("A1").Resize(1, 4).Value = Array(H_ID, H_DATE, H_TOPIC, H_RULE)
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:D2"), , xlYes)
        lo.Name = PROJ_TABLE
    End If

    EnsureColumn lo, H_ID
    cScope = EnsureColumn(lo, H_SCOPE)
    EnsureColumn lo, H_ORIGIN
    cRule = ColIdx(lo, H_RULE)
    If cRule = 0 Then
        MsgBox UX("tblRules thi\u1EBFu c\u1ED9t """) & H_RULE & UX(""". \u0110\u1ED5i t\u00EAn c\u1ED9t n\u1ED9i dung quy t\u1EAFc r\u1ED3i ch\u1EA1y l\u1EA1i."), vbExclamation
        Exit Sub
    End If

    If Not lo.DataBodyRange Is Nothing Then
        With lo.ListColumns(cScope).DataBodyRange
            .Validation.Delete
            .Validation.Add Type:=xlValidateList, Formula1:=SCOPE_SHARED & "," & SCOPE_LOCAL
            For Each cell In .Cells
                If Trim$(CStr(cell.Value)) = "" And Trim$(CStr(lo.Range.Cells(cell.Row - lo.Range.Row + 1, cRule).Value)) <> "" Then
                    cell.Value = SCOPE_LOCAL
                End If
            Next cell
        End With
    End If

    CreateMasterFile
    AddMasterQuery wb
    MsgBox "OK: tblRules + " & VIEW_SHEET & " (" & MASTER_PATH & ")", vbInformation
End Sub

'---------- Buoc 3: query keo Master -> sheet _RULES_MASTER ----------
Public Sub AddMasterQuery(wb As Workbook)
    Dim fParam As String, fMain As String, ws As Worksheet, lo As ListObject
    fParam = """" & MASTER_PATH & """ meta [IsParameterQuery=true, Type=""Text"", IsParameterQueryRequired=true]"
    fMain = "let" & vbCrLf & _
        "    Src = Excel.Workbook(File.Contents(pMasterPath), true)," & vbCrLf & _
        "    Tbl = Src{[Item=""" & MASTER_TABLE & """, Kind=""Table""]}[Data]," & vbCrLf & _
        "    StatusCol = Table.ColumnNames(Tbl){5}," & vbCrLf & _
        "    Kept = Table.SelectRows(Tbl, each Text.Upper(Text.Trim(Text.From(Record.Field(_, StatusCol) ?? """"))) <> ""OBSOLETE"" and Record.Field(_, Table.ColumnNames(Tbl){0}) <> null)" & vbCrLf & _
        "in" & vbCrLf & _
        "    Kept"
    UpsertQuery wb, "pMasterPath", fParam
    UpsertQuery wb, "qRulesMaster", fMain

    On Error Resume Next
    Set ws = wb.Worksheets(VIEW_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = VIEW_SHEET
        ws.Tab.Color = RGB(166, 166, 166)
    End If
    Set lo = FindTable(wb, VIEW_TABLE)
    If lo Is Nothing Then
        ws.Cells.Clear
        Set lo = ws.ListObjects.Add(SourceType:=0, _
            Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=qRulesMaster;Extended Properties=""""", _
            Destination:=ws.Range("A1"))
        With lo.QueryTable
            .CommandType = xlCmdSql
            .CommandText = Array("SELECT * FROM [qRulesMaster]")
            .RefreshOnFileOpen = True
            .BackgroundQuery = False
            .Refresh BackgroundQuery:=False
        End With
        lo.Name = VIEW_TABLE
    Else
        lo.QueryTable.Refresh BackgroundQuery:=False
    End If
End Sub

Private Sub UpsertQuery(wb As Workbook, ByVal qName As String, ByVal formula As String)
    Dim q As Object
    On Error Resume Next
    Set q = wb.Queries(qName)
    On Error GoTo 0
    If q Is Nothing Then
        wb.Queries.Add Name:=qName, Formula:=formula
    Else
        q.Formula = formula
    End If
End Sub

'---------- Buoc 4: day dong CHUNG chua co ID vao Master ----------
Public Sub PushRulesNow()
    PushSharedRules ActiveWorkbook, False
End Sub

Public Sub PushSharedRules(wb As Workbook, Optional ByVal silent As Boolean = True)
    Dim lo As ListObject, mlo As ListObject, mwb As Workbook, openedByUs As Boolean
    Dim cId As Long, cScope As Long, cRule As Long, cTopic As Long, cDt As Long, cOrigin As Long
    Dim r As Long, n As Long, nextNum As Long, newId As String, ruleTxt As String, origin As String
    Dim rw As Range, dst As Range, oldEvents As Boolean, oldScreen As Boolean

    If gBusy Then Exit Sub
    If StrComp(wb.FullName, MASTER_PATH, vbTextCompare) = 0 Then Exit Sub
    Set lo = FindTable(wb, PROJ_TABLE)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    cId = ColIdx(lo, H_ID): cScope = ColIdx(lo, H_SCOPE): cRule = ColIdx(lo, H_RULE)
    If cId = 0 Or cScope = 0 Or cRule = 0 Then Exit Sub   ' file chua Setup
    cTopic = ColIdx(lo, H_TOPIC): cDt = ColIdx(lo, H_DATE): cOrigin = ColIdx(lo, H_ORIGIN)

    For r = 1 To lo.ListRows.Count
        Set rw = lo.ListRows(r).Range
        If UCase$(Trim$(CStr(rw.Cells(1, cScope).Value))) = SCOPE_SHARED _
           And Trim$(CStr(rw.Cells(1, cId).Value)) = "" _
           And Trim$(CStr(rw.Cells(1, cRule).Value)) <> "" Then n = n + 1
    Next r
    If n = 0 Then
        If Not silent Then MsgBox UX("Kh\u00F4ng c\u00F3 quy t\u1EAFc CHUNG m\u1EDBi."), vbInformation
        Exit Sub
    End If

    gBusy = True
    oldEvents = Application.EnableEvents: oldScreen = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    On Error GoTo Fail

    CreateMasterFile
    On Error Resume Next
    Set mwb = Workbooks(Mid$(MASTER_PATH, InStrRev(MASTER_PATH, "\") + 1))
    On Error GoTo Fail
    If mwb Is Nothing Then
        Set mwb = Workbooks.Open(MASTER_PATH, UpdateLinks:=0)
        openedByUs = True
        wb.Activate
    End If
    If mwb.ReadOnly Then
        If openedByUs Then mwb.Close False
        Err.Raise vbObjectError + 1, , "Master " & UX("\u0111ang b\u1ECB kho\u00E1 (read-only). S\u1EBD \u0111\u1EA9y l\u1EA1i \u1EDF l\u1EA7n Save sau.")
    End If
    Set mlo = FindTable(mwb, MASTER_TABLE)
    If mlo Is Nothing Then Err.Raise vbObjectError + 2, , UX("Kh\u00F4ng t\u00ECm th\u1EA5y ") & MASTER_TABLE

    nextNum = MaxIdNum(mlo) + 1
    origin = wb.Name
    If InStrRev(origin, ".") > 0 Then origin = Left$(origin, InStrRev(origin, ".") - 1)
    n = 0

    For r = 1 To lo.ListRows.Count
        Set rw = lo.ListRows(r).Range
        ruleTxt = Trim$(CStr(rw.Cells(1, cRule).Value))
        If UCase$(Trim$(CStr(rw.Cells(1, cScope).Value))) = SCOPE_SHARED _
           And Trim$(CStr(rw.Cells(1, cId).Value)) = "" And ruleTxt <> "" Then
            newId = FindIdByRule(mlo, ruleTxt)          ' trung noi dung -> dung lai ID
            If newId = "" Then
                newId = "R" & Format$(nextNum, "0000"): nextNum = nextNum + 1
                Set dst = NewRow(mlo)
                SetCell mlo, dst, H_ID, newId
                If cDt > 0 And Trim$(CStr(rw.Cells(1, cDt).Value)) <> "" Then
                    SetCell mlo, dst, H_DATE, rw.Cells(1, cDt).Value
                Else
                    SetCell mlo, dst, H_DATE, Date
                End If
                If cTopic > 0 Then SetCell mlo, dst, H_TOPIC, rw.Cells(1, cTopic).Value
                SetCell mlo, dst, H_RULE, ruleTxt
                SetCell mlo, dst, H_ORIGIN, origin
                SetCell mlo, dst, H_STATUS, STATUS_NEW
                SetCell mlo, dst, H_VER, 1
                n = n + 1
            End If
            rw.Cells(1, cId).Value = newId
            If cOrigin > 0 Then
                If Trim$(CStr(rw.Cells(1, cOrigin).Value)) = "" Then rw.Cells(1, cOrigin).Value = origin
            End If
        End If
    Next r

    mwb.Save
    If openedByUs Then mwb.Close False
    RefreshView wb

Done:
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    gBusy = False
    If Not silent Then MsgBox UX("\u0110\u00E3 \u0111\u1EA9y ") & n & UX(" quy t\u1EAFc v\u00E0o Master."), vbInformation
    Exit Sub
Fail:
    Application.EnableEvents = oldEvents: Application.ScreenUpdating = oldScreen
    gBusy = False
    MsgBox UX("L\u1ED7i \u0111\u1ED3ng b\u1ED9 Master: ") & Err.Description, vbExclamation
End Sub

Private Sub SetCell(mlo As ListObject, dst As Range, ByVal header As String, ByVal v As Variant)
    Dim c As Long
    c = ColIdx(mlo, header)
    If c > 0 Then dst.Cells(1, c).Value = v
End Sub

Private Function MaxIdNum(mlo As ListObject) As Long
    Dim c As Long, cell As Range, s As String
    c = ColIdx(mlo, H_ID)
    If c = 0 Or mlo.DataBodyRange Is Nothing Then Exit Function
    For Each cell In mlo.ListColumns(c).DataBodyRange.Cells
        s = Trim$(CStr(cell.Value))
        If UCase$(Left$(s, 1)) = "R" And IsNumeric(Mid$(s, 2)) Then
            If CLng(Mid$(s, 2)) > MaxIdNum Then MaxIdNum = CLng(Mid$(s, 2))
        End If
    Next cell
End Function

Private Function FindIdByRule(mlo As ListObject, ByVal ruleTxt As String) As String
    Dim cR As Long, cI As Long, r As Long
    cR = ColIdx(mlo, H_RULE): cI = ColIdx(mlo, H_ID)
    If cR = 0 Or cI = 0 Or mlo.DataBodyRange Is Nothing Then Exit Function
    For r = 1 To mlo.ListRows.Count
        If StrComp(Trim$(CStr(mlo.ListRows(r).Range.Cells(1, cR).Value)), ruleTxt, vbTextCompare) = 0 Then
            FindIdByRule = CStr(mlo.ListRows(r).Range.Cells(1, cI).Value): Exit Function
        End If
    Next r
End Function

Private Sub RefreshView(wb As Workbook)
    Dim lo As ListObject
    On Error Resume Next
    Set lo = FindTable(wb, VIEW_TABLE)
    If Not lo Is Nothing Then lo.QueryTable.Refresh BackgroundQuery:=False
End Sub
