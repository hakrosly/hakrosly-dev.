"""vba/ のモジュールを LibreOffice (VBA 互換モード) に読み込んで、自己診断を実行する。
   使い方: python3 run_tests.py
   これは開発する側の確かめ用。会社の PC では不要 (Excel では「自己診断」ボタンを使う)。"""
import os, sys, shutil, traceback, openpyxl
import lo

here = os.path.dirname(os.path.abspath(__file__))
work = os.path.join(here, "work")

DRV = """
Option Explicit
Function Drv_Setup() As String
    On Error GoTo EH
    SetupTool
    Drv_Setup = "OK"
    Exit Function
EH:
    Drv_Setup = "ERR " & Err.Number & " " & Err.Description
End Function
Function Drv_Test() As String
    On Error GoTo EH
    Drv_Test = RunSelfTest()
    Exit Function
EH:
    Drv_Test = "ERR " & Err.Number & " " & Err.Description & " / after " & TestProgress()
End Function
"""

def sheet_rows(doc, name, ncol, maxrow=400):
    sh = doc.Sheets.getByName(name)
    rows = []
    for r in range(maxrow):
        row = [sh.getCellByPosition(c, r).getString() for c in range(ncol)]
        if not any(row):
            break
        rows.append(row)
    return rows

def main():
    shutil.rmtree(work, ignore_errors=True)
    os.makedirs(work)
    os.chdir(work)
    wb = openpyxl.Workbook()
    wb.active.title = "Sheet1"
    wb.save("tool.xlsx")
    p, ctx = lo.start()
    ok = False
    try:
        doc = lo.load(ctx, "tool.xlsx")
        mods = lo.tool_modules(os.path.join(here, "..", "vba"), os.path.join(here, "lo_shim"))
        lo.add_modules(doc, mods + [("modDrv", DRV, "std")])
        print("Setup:", lo.run(doc, "modDrv.Drv_Setup"))
        print("RunSelfTest:", lo.run(doc, "modDrv.Drv_Test"))
        rows = sheet_rows(doc, "テスト結果", 5)
        bad = 0
        for r in rows[1:]:
            mark = "ok " if r[4] == "合格" else "NG "
            if r[4] != "合格":
                bad += 1
                print(mark, r[0], r[1], "\n     期待:", r[2].replace("\n", "\\n")[:300], "\n     実際:", r[3].replace("\n", "\\n")[:300])
            else:
                print(mark, r[0], r[1])
        print(f"{len(rows)-1} tests, {bad} failed")
        ok = (bad == 0 and len(rows) > 1)
    except Exception:
        traceback.print_exc()
    finally:
        lo.stop(p)
    sys.exit(0 if ok else 1)

if __name__ == "__main__":
    main()
