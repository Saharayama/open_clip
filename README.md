# Open Clipboard
## ショートカットの作成
以下のようにショートカットのリンク先を設定する．`python.exe`ではなく`pythonw.exe`を使用することで実行時のコンソール画面が非表示となる．
```bat
%HOMEPATH%\open_clip\.venv\Scripts\pythonw.exe %HOMEPATH%\open_clip\open_clip.py
```
## PyInstallerでexe化
```sh
pyinstaller.exe open_clip.py --onefile --noconsole
```
