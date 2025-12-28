uv synv
uv add pyinstaller
cd .venv\Scripts
.\activate
pyinstaller main.py --add-data index.html:.
