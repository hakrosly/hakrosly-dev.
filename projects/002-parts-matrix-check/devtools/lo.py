# LibreOffice を裏で動かして、VBA 互換モードでマクロを実行するための道具
import subprocess, time, os, sys, uno
from com.sun.star.beans import PropertyValue

PORT = 2002
def start():
    prof = os.path.abspath("lo_profile")
    p = subprocess.Popen(["soffice", "--headless", "--invisible", "--norestore", "--nologo",
        f"-env:UserInstallation=file://{prof}",
        f"--accept=socket,host=localhost,port={PORT};urp;"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    local = uno.getComponentContext()
    resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
    for i in range(60):
        try:
            ctx = resolver.resolve(f"uno:socket,host=localhost,port={PORT};urp;StarOffice.ComponentContext")
            return p, ctx
        except Exception:
            time.sleep(0.5)
    raise RuntimeError("soffice did not start")

def pv(n, v):
    x = PropertyValue(); x.Name = n; x.Value = v; return x

def load(ctx, path, hidden=True):
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    return desktop.loadComponentFromURL(uno.systemPathToFileUrl(os.path.abspath(path)), "_blank", 0,
        (pv("Hidden", hidden), pv("MacroExecutionMode", 4)))

def add_modules(doc, modules):
    libs = doc.BasicLibraries
    libs.VBACompatibilityMode = True
    if not libs.hasByName("Standard"):
        libs.createLibrary("Standard")
    lib = libs.getByName("Standard")
    for name, src, kind in modules:
        head = "Rem Attribute VBA_ModuleType=VBAModule\nOption VBASupport 1\n"
        if kind == "class":
            head = "Rem Attribute VBA_ModuleType=VBAClassModule\nOption VBASupport 1\nOption ClassModule\n"
        if lib.hasByName(name):
            lib.removeByName(name)
        lib.insertByName(name, head + src)

def run(doc, macro, *args):
    sp = doc.getScriptProvider()
    s = sp.getScript(f"vnd.sun.star.script:Standard.{macro}?language=Basic&location=document")
    return s.invoke(tuple(args), (), ())[0]

def read(path):
    with open(path, encoding="utf-8") as f:
        return f.read()

def tool_modules(vba_dir, shim_dir):
    """vba/*.bas を読む。modDict だけ LibreOffice 用に差し替える"""
    mods = [("clsBox", read(os.path.join(shim_dir, "clsBox.cls")), "class"),
            ("clsDict", read(os.path.join(shim_dir, "clsDict.cls")), "class")]
    for fn in sorted(os.listdir(vba_dir)):
        if not fn.endswith(".bas"):
            continue
        name = fn[:-4]
        src = read(os.path.join(shim_dir, "modDict.bas")) if name == "modDict" else read(os.path.join(vba_dir, fn))
        mods.append((name, src, "std"))
    return mods

def stop(p):
    p.terminate()
    try:
        p.wait(timeout=20)
    except Exception:
        p.kill()
