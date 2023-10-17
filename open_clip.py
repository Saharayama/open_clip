import os
import re
import subprocess

import pyperclip


def format_path(p: str) -> str:
    path: str = ""
    p = re.sub(r'"|^ *-|^\t*-', "", p).strip()

    if ":" in p and "/" in p:
        path = p.replace("/", "\\")
    elif "/" in p:
        path = r"\\" + p.replace("/", "\\")
    elif r"\\" in p:
        p = re.sub("^\\\\+", "", p)
        path = rf"\\{p}"
    else:
        path = p

    path = rf'"{path}"'

    return path


def open_path(path_formatted: str) -> None:
    path_suffix: str = os.path.splitext(path_formatted)[1]
    if "-->" in path_suffix:
        return

    office_path: str = r"C:\Program Files\Microsoft Office\root\Office16"
    cmd: str | None = None

    if "xls" in path_suffix:
        cmd = rf'"{office_path}\EXCEL.EXE" /x {path_formatted}'
    elif "doc" in path_suffix:
        cmd = rf'"{office_path}\WINWORD.EXE" /w {path_formatted}'

    if cmd is None:
        os.startfile(path_formatted)
    else:
        subprocess.Popen(cmd)


cb: str = pyperclip.paste()

for line in cb.splitlines():
    if not line.strip():
        continue
    try:
        open_path(format_path(line))
    except FileNotFoundError:
        continue
