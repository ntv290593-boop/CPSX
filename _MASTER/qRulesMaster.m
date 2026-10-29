// Query 1 — tên: pMasterPath  (tham số đường dẫn)
"D:\COSTING EXCEL\_MASTER\RULES_MASTER.xlsx" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]

// Query 2 — tên: qRulesMaster  (Load To > Table > sheet _RULES_MASTER, tên bảng tblRulesMasterView)
let
    Src = Excel.Workbook(File.Contents(pMasterPath), true),
    Tbl = Src{[Item="tblRulesMaster", Kind="Table"]}[Data],
    StatusCol = Table.ColumnNames(Tbl){5},
    Kept = Table.SelectRows(Tbl, each Text.Upper(Text.Trim(Text.From(Record.Field(_, StatusCol) ?? ""))) <> "OBSOLETE" and Record.Field(_, Table.ColumnNames(Tbl){0}) <> null)
in
    Kept
