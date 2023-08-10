import pyperclip
import subprocess
import os
import re

EXCEL_SUFFIX_LIST: list[str] = ['.xls"', '.xlsx"']
WORD_SUFFIX_LIST: list[str] = ['.doc"', '.docx"']
VSCODE_SUFFIX_LIST: list[str] = ['.md"', '.c"', '.txt"', '.json"', '.h"']


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
    if path_suffix in EXCEL_SUFFIX_LIST:
        cmd: str = (
            r'"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" /x {}'.format(
                path_formatted
            )
        )
        subprocess.Popen(cmd)
    elif path_suffix in WORD_SUFFIX_LIST:
        cmd: str = r'"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE" /w {}'.format(
            path_formatted
        )
        subprocess.Popen(cmd)
    elif path_suffix in VSCODE_SUFFIX_LIST:
        cmd: str = r'"C:\Program Files\Microsoft VS Code\Code.exe" {}'.format(
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
