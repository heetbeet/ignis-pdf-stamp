1>2# : ^
'''
@echo off
"%~dp0\\..\\bin\\python\\python.exe" "%~dpf0" %*
exit /b
'''

from pathlib import Path
import subprocess
import tempfile

this_dir = Path(__file__).resolve().parent
installbase = Path(__file__).resolve().parent.parent

sanitize = str(installbase).replace("/","\\").replace("\\","\\\\")

reg = rf"""Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Classes\*\shell\apply_ignis_stamp]
@="Apply Ignis Stamp"
"icon"="{sanitize}\\assets\\IgnisStamp.svg.ico"

[HKEY_CURRENT_USER\Software\Classes\*\shell\apply_ignis_stamp\command]
@="\"{sanitize}\\scripts\\ignis-stamp.cmd\" \"%1\""
"""

with tempfile.TemporaryDirectory() as td:
    td = Path(td)
    with td.joinpath("add-to-menu.reg").open("w") as f:
        f.write(reg)

    subprocess.call(["reg", "import", str(td.joinpath("add-to-menu.reg"))], stderr=subprocess.PIPE)