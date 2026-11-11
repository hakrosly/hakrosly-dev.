import os, lo, openpyxl, traceback
here = os.path.dirname(os.path.abspath(__file__))
os.makedirs(os.path.join(here, "work"), exist_ok=True); os.chdir(os.path.join(here, "work"))
wb = openpyxl.Workbook(); wb.save("b.xlsx")
mods = lo.tool_modules(os.path.join(here, "..", "vba"), os.path.join(here, "lo_shim"))
order = ["clsDict", "modConst", "modDict", "modUtil", "modText", "modLog", "modFile", "modBom", "modMatrix", "modSpecMap", "modCompare", "modReport", "modRun", "modSetup", "modTestData", "modTest", "modMain"]
mods.sort(key=lambda m: order.index(m[0]))
ping = 'Option Explicit\nFunction Ping() As String\n    Ping = "pong"\nEnd Function\n'
p, ctx = lo.start()
try:
    doc = lo.load(ctx, "b.xlsx")
    for i in range(len(mods)):
        lo.add_modules(doc, [mods[i], ("modPing", ping, "std")])
        try:
            r = lo.run(doc, "modPing.Ping")
        except Exception as e:
            r = "EXC " + str(e)[:300]
        print(mods[i][0], "->", r)
finally:
    lo.stop(p)
