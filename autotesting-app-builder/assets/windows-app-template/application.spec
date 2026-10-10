# -*- mode: python ; coding: utf-8 -*-
import os
from pathlib import Path
from PyInstaller.utils.hooks import collect_dynamic_libs
ROOT=Path(SPECPATH)
APP_NAME=os.environ.get("AUTOTESTING_APP_NAME","ScoringApp")
QT=collect_dynamic_libs("PySide6",search_patterns=["Qt6Core.dll","Qt6Gui.dll","Qt6Widgets.dll","Qt6Network.dll","pyside6.abi3.dll","plugins/platforms/qwindows.dll"])
# config/ is required: project.json is read by the app and by the sources loop below.
datas=[(str(ROOT/"config"),"config")]
# templates/ holds per-project artifacts, so a project may not have created it yet
# -- and may keep its templates elsewhere entirely. The configured template files
# are packed by the sources loop below, so a missing directory must not abort the
# build. PyInstaller fails hard on a datas source that does not exist.
for directory in ("templates","scripts"):
    if (ROOT/directory).is_dir(): datas.append((str(ROOT/directory),directory))
datas.append((str(ROOT/"src/ui_design_system/workbench.qss"),"src/ui_design_system"))
if (ROOT/"assets").is_dir(): datas.append((str(ROOT/"assets"),"assets"))
import json
for source in json.loads((ROOT/'config/project.json').read_text(encoding='utf-8'))['sources']:
    rel=Path(source['path'])
    datas.append((str(ROOT/rel),rel.parent.as_posix()))
a=Analysis([str(ROOT/"src"/"main.py")],pathex=[str(ROOT/"src")],binaries=QT,datas=datas,hiddenimports=[],hookspath=[],hooksconfig={},runtime_hooks=[],excludes=[],noarchive=False,optimize=0)
pyz=PYZ(a.pure)
exe=EXE(pyz,a.scripts,[],exclude_binaries=True,name=APP_NAME,debug=False,bootloader_ignore_signals=False,strip=False,upx=False,console=False,disable_windowed_traceback=False,argv_emulation=False)
coll=COLLECT(exe,a.binaries,a.datas,strip=False,upx=False,upx_exclude=[],name=APP_NAME)
