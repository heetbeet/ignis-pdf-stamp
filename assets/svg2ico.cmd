@staticmethod #  2>nul &@echo off
def _():''' 2>nul
    ::*******************************************************
    :: This is a .cmd header to define how to run this file
    ::*******************************************************
    setlocal
    set "PATH=%~dp0..\bin\python;%userprofile%\Applications\image_magic;%PATH%"
    python "%~dp0%~n0.cmd"
    exit /b %errorlevel%
'''

import subprocess
import tempfile
import locate
from pathlib import Path
import os
import shutil

if __name__ == "__main__":
    
    tmp = Path(tempfile.mktemp())
    os.makedirs(tmp, exist_ok=True)
    
    # https://stackoverflow.com/a/9035861/1490584
    convert_exe = [i.strip() for i in subprocess.check_output(["where", "convert"]).decode("utf-8").split("\n") if i.strip()!=""]
    if len(convert_exe) == 0:
        raise FileNotFoundError("convert.exe")
    convert_exe = convert_exe[0]
    
    
    for i in locate.this_dir().glob("*.svg"):
        
        
        sizes = [16, 32, 40, 48, 128, 256]
        for j in sizes:
            cmd = ["inkscape", "--export-type=png", "-w", str(j), "-h", str(j), str(i), rf"--export-filename={tmp}\{j}.png"]
            subprocess.call(cmd)
            
        files = [rf"{tmp}\{j}.png" for j in sizes]
        subprocess.call([convert_exe] + files + [str(i.parent.joinpath(i.name+".ico"))])


    shutil.rmtree(tmp, ignore_errors=True)