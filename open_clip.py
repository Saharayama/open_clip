import pyperclip
import subprocess
import os
import re

def format_path(p: str) -> str:
    path: str = ""
    p = re.sub('"|^-', "", p)
    p = p.strip()

    if "/" in p:
        path = r"\\" + re.sub("/", r"\\", p)
    elif "\\\\" in p:
        p = re.sub("^\\\\+", "", p)
        path = r"\\" + p
    else:
        path = p

    path = '"' + path + '"'

    return path


def open_path(path_formatted: str) -> None:
    path_suffix: str = os.path.splitext(path_formatted)[1]
    if "xls" in path_suffix:
        cmd: str = (
            r'"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" /x {}'.format(
                path_formatted
            )
        )
        subprocess.Popen(cmd)
    elif "doc" in path_suffix:
        cmd: str = r'"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE" /w {}'.format(
            path_formatted
        )
        subprocess.Popen(cmd)
    else:
        os.startfile(path_formatted)


cb: str = pyperclip.paste()

for line in cb.splitlines():
    if line == "" or line.isspace():
        continue
    else:
        try:
            open_path(format_path(line))
        except FileNotFoundError:
            continue
