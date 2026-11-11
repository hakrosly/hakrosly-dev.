"""VBA の簡単な点検。LibreOffice では動かせない所 (画面を出す所など) の書き間違いを見つけるため。
   - 宣言していない名前を使っていないか
   - オブジェクトを Set なしで代入していないか (よくある形だけ)
   - 外部通信などの禁止された言葉が無いか (構想書の T-41)
   - CP932 (Shift-JIS) で保存できるか、1 行が長すぎないか
   使い方: python3 lint_vba.py"""
import os, re, sys
here = os.path.dirname(os.path.abspath(__file__))
vba = os.path.join(here, "..", "vba")
KEYWORDS = set("""and as boolean byref byval call case const cstr clng cdbl cbool dim do each else elseif empty end error exit false for function get goto if in is let like long loop mod new next not nothing null object on option or private property public redim resume select set step string sub then to true type until variant wend while with xor explicit preserve date double integer me optional""".split())
BUILTINS = set("""abs array ascw asc chr chrw cstr clng cdbl cbool cint cdate day dir environ err filedatetime filelen format hour iif inputbox instr instrrev int isarray isdate isempty iserror ismissing isnull isnumeric isobject join lbound lcase left len ltrim mid minute mkdir month msgbox now replace right rnd rtrim second space split str strcomp strconv string trim typename ubound ucase val vartype year createobject
application thisworkbook workbooks worksheets activeworkbook activesheet range cells rows columns intersect union debug
vbcr vblf vbcrlf vbtab vbnarrow vbwide vbstring vbboolean vbdouble vbdate vbbinarycompare vbtextcompare vbdirectory vbyesno vbyes vbno vbok vbokcancel vbquestion vbinformation vbexclamation vbcritical
workbook worksheet collection""".split())
FORBIDDEN = ["winhttp", "xmlhttp", "shell", "declare ", "internetexplorer", "urldownload", "querytables", "msxml", "wscript", "sendkeys", "createobject(\"wscript", "kill "]

def strip(line):
    out = []; q = False
    for ch in line:
        if ch == '"': q = not q; out.append('"'); continue
        if not q and ch == "'": break
        out.append(" " if q else ch)
    return "".join(out)

mods = {}
for fn in sorted(os.listdir(vba)):
    if fn.endswith(".bas"):
        mods[fn[:-4]] = open(os.path.join(vba, fn), encoding="utf-8").read()

problems = []
glob = set()
for name, src in mods.items():
    try: src.encode("cp932")
    except UnicodeEncodeError as e: problems.append(f"{name}: CP932 で保存できない文字 {src[e.start:e.end]!r}")
    for i, raw in enumerate(src.split("\n"), 1):
        if len(raw) > 1000: problems.append(f"{name}:{i}: 1 行が長すぎる ({len(raw)})")
        low = strip(raw).lower()
        for w in FORBIDDEN:
            if w in low and not (w == "shell" and False):
                problems.append(f"{name}:{i}: 禁止された言葉 {w.strip()!r}")
        m = re.match(r"\s*(?:Public |Private )?(?:Sub|Function|Property Get|Property Let) (\w+)", raw)
        if m: glob.add(m.group(1).lower())
        m = re.match(r"\s*(?:Public |Private )?Const (\w+)", raw)
        if m: glob.add(m.group(1).lower())
        if re.match(r"(Public|Private) (?!Sub|Function|Const|Type|Property)", raw):
            for v in re.findall(r"(\w+)(?:\(\))? As ", raw): glob.add(v.lower())

def procs(src):
    cur = None; body = []
    for raw in src.split("\n"):
        if re.match(r"\s*(?:Public |Private )?(?:Sub|Function) \w+", raw) and cur is None:
            cur = raw; body = [raw]
        elif cur is not None:
            body.append(raw)
            if re.match(r"\s*End (Sub|Function)", raw):
                yield body; cur = None

for name, src in mods.items():
    for body in procs(src):
        text = []; cont = ""
        for raw in body:
            s = strip(raw).rstrip()
            if s.endswith(" _"): cont += s[:-1]; continue
            text.append(cont + s); cont = ""
        local = set()
        head = text[0]
        pname = re.search(r"(?:Sub|Function) (\w+)", head).group(1)
        local.add(pname.lower())
        for v in re.findall(r"(?:ByVal|ByRef) (\w+)", head): local.add(v.lower())
        labels = set()
        for s in text:
            if re.match(r"\s*(Dim|ReDim)\b", s):
                for v in re.findall(r"(?:Dim |, |ReDim |Preserve )(\w+)(?:\([^)]*\))?(?= As |,|$|\()", s): local.add(v.lower())
            m = re.match(r"^(\w+):\s*$", s)
            if m: labels.add(m.group(1).lower())
            m = re.match(r"\s*For Each (\w+)", s)
        for ln, s in enumerate(text[1:], 1):
            if re.match(r"\s*Dim\b", s): continue
            s2 = re.sub(r"\.\w+", ".", s)                 # メンバー (obj.Name) は見ない
            s2 = re.sub(r"\b\w+:=", "", s2)               # 名前付き引数
            s2 = re.sub(r"&H[0-9A-Fa-f]+", "0", s2)
            for w in re.findall(r"[A-Za-z_]\w*", s2):
                wl = w.lower()
                if wl in KEYWORDS or wl in BUILTINS or wl in local or wl in glob or wl in labels: continue
                problems.append(f"{name}.{pname}: 宣言の無い名前 {w!r}  |  {s.strip()[:90]}")
            # Set なしのオブジェクト代入 (よくある形だけ)
            if ".Address(" not in s and re.match(r"\s*(?!Set )\w+ = (NewDict\(\)|\w+\.(Range|Cells|Worksheets)\([^=]*\)$|Workbooks\.(Add|Open)|New )", s):
                problems.append(f"{name}.{pname}: Set が要りそう  |  {s.strip()[:90]}")

seen = set()
for p in problems:
    if p not in seen:
        seen.add(p); print(p)
print(f"{len(mods)} modules, {sum(len(s.splitlines()) for s in mods.values())} lines, {len(seen)} problems")
sys.exit(1 if seen else 0)
