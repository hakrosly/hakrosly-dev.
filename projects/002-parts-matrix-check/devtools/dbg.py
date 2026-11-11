import os, sys, lo, openpyxl, traceback
here = os.path.dirname(os.path.abspath(__file__))
os.makedirs(os.path.join(here, "work"), exist_ok=True); os.chdir(os.path.join(here, "work"))
wb = openpyxl.Workbook(); wb.save("d.xlsx")
body = sys.stdin.read()
drv = 'Option Explicit\nFunction Dbg() As String\n    Dim s As String\n    On Error GoTo EH\n' + body + '\n    Dbg = s & " OK"\n    Exit Function\nEH:\n    Dbg = s & " ERR " & Err.Number & " " & Err.Description\nEnd Function\n'
p, ctx = lo.start()
try:
    doc = lo.load(ctx, "d.xlsx")
    lo.add_modules(doc, lo.tool_modules(os.path.join(here, "..", "vba"), os.path.join(here, "lo_shim")) + [("modDbg", drv, "std")])
    print(lo.run(doc, "modDbg.Dbg"))
finally:
    lo.stop(p)
