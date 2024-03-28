# Open Clipboard
## ショートカットで使う
以下のようにショートカットのリンク先を設定する．`python.exe`ではなく`pythonw.exe`を使用することで実行時のコンソール画面が非表示となる．
```bat
D:\myapps\script\open_clip\.venv\Scripts\pythonw.exe D:\myapps\script\open_clip\open_clip.py
```
## PyInstaller
```sh
pyinstaller.exe open_clip.py --onefile --noconsole
```
