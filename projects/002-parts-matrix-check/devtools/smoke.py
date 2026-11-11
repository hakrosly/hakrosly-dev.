import openpyxl, lo, os, traceback
here = os.path.dirname(os.path.abspath(__file__))
os.makedirs(os.path.join(here, "work"), exist_ok=True)
os.chdir(os.path.join(here, "work"))
wb = openpyxl.Workbook(); ws = wb.active; ws.title = "M"
ws["A1"] = "title"
ws.merge_cells("C3:D3"); ws["C3"] = "STD"; ws.merge_cells("E3:F3"); ws["E3"] = "DX"
for i, v in enumerate(["2WD", "4WD", "2WD", "4WD"]): ws.cell(4, 3 + i, v)
rows = [("ブラケットA1", "AB123-0010", "○", "○", None, None), ("ブラケットB", "AB123-0020", None, None, "○", None),
        ("ブラケットＣ", "ａｂ１２３－００３０", "○", "○", "○", "○"), ("ブラケットD", "AB123-0040", "○", "○", None, "○"),
        ("ブラケットF", "AB123-0060", "○", None, None, None)]
for r, row in enumerate(rows):
    for c, v in enumerate(row): ws.cell(5 + r, 1 + c, v)
b = wb.create_sheet("B")
b.append(["品番", "部品名称", "仕様コード", "員数"])
for part, name, specs, q in [("AB123-0010", "ブラケットA", "S1,S2", 1), ("AB123-0020", "ブラケットB", "S3,S4", 1), ("AB123-0030", "ブラケットC", "S1,S2,S3,S4", 2),
                             ("AB123-0040", "ブラケットD", "S2,S4", 1), ("AB123-0050", "ブラケットE", "S3", 1), ("ZZ999-0001", "カバー", "S1", 1)]:
    for s in specs.split(","): b.append([part, name, s, q])
wb.save("smoke.xlsx")
drv = r'''
Option Explicit
Function Main() As String
    Dim s As String, lay As Object, info As Object, a As Variant, i As Long, cfg As Object, w As Variant
    On Error GoTo EH
    TextDefaults
    s = s & NormPart("ａｂ１２３－００１０") & "|" & NormPart(" ab123-0010 ") & "|" & NormPart("AB123" & ChrW(&H2010) & "0010") & "|" & NormName("ブラケットＣ") & vbLf
    LogReset
    Set lay = NewLayout("t", "M", O_VERT, "B5:B9", "A5:A9", "C3:F4", K_MARK, OPT_YES, OPT_NOSPLIT, "")
    Set info = NewDict()
    a = ReadMatrixSheet(ThisWorkbook.Worksheets("M"), lay, info)
    s = s & "parts=" & info.Item("parts") & " cols=" & info.Item("cols") & " cells=" & info.Item("cells") & " yes=" & info.Item("yes") & " raw=" & info.Item("rawYes") & " used=" & info.Item("usedYes") & vbLf
    s = s & "heads=" & Replace(info.Item("headings"), vbLf, ";") & vbLf
    For i = 1 To 5
        s = s & a(i, MC_KEY) & "," & a(i, MC_NAME) & "," & a(i, MC_HEAD) & "," & a(i, MC_VAL) & "," & a(i, MC_VALCELL) & "," & a(i, MC_HEADCELL) & vbLf
    Next i
    Set cfg = NewBomConfig("", 1, "品番", "部品名称", "仕様コード", "", "員数", "品番", "AB123*")
    Set info = NewDict()
    a = ReadBomSheet(ThisWorkbook.Worksheets("B"), cfg, info)
    s = s & "bom rows=" & RowCount(a) & " read=" & info.Item("readRows") & " countA=" & info.Item("countA") & " data=" & info.Item("dataRows") & vbLf
    s = s & a(1, BC_KEY) & "," & a(1, BC_SPECKEY) & "," & a(1, BC_QTY) & "," & a(1, BC_INSCOPE) & "," & a(12, BC_KEY) & "," & a(12, BC_INSCOPE) & "," & a(12, BC_ROW) & vbLf
    s = s & "warn=" & LogCount() & vbLf
    Main = s & "OK"
    Exit Function
EH:
    Main = s & "ERR " & Err.Number & " " & Err.Description
End Function
'''
p, ctx = lo.start()
try:
    doc = lo.load(ctx, "smoke.xlsx")
    mods = lo.tool_modules(os.path.join(here, "..", "vba"), os.path.join(here, "lo_shim"))
    lo.add_modules(doc, mods + [("modDrv", drv, "std")])
    print(lo.run(doc, "modDrv.Main"))
except Exception:
    traceback.print_exc()
finally:
    lo.stop(p)
