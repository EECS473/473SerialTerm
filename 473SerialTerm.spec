# 473SerialTerm.spec — low false-positive build
from pathlib import Path
import PySide6
from PyInstaller.utils.hooks import collect_submodules

HERE = Path(__file__).parent if "__file__" in globals() else Path.cwd()

ICO = HERE / "assets" / "473SerialTermIcon.ico"
if not ICO.exists():
    raise SystemExit(f"Icon not found: {ICO}. Place it there or fix the path.")

# Version resource (create file_version_info.txt next to this spec)
VER = HERE / "file_version_info.txt"

base = Path(PySide6.__file__).parent
plugins = base / "plugins"

qt_plugin_files = [
    plugins / "platforms" / "qwindows.dll",
    plugins / "imageformats" / "qico.dll",
    plugins / "imageformats" / "qjpeg.dll",
]

datas = [(str(p), f"PySide6/plugins/{p.parent.name}") for p in qt_plugin_files]
datas += [(str(ICO), 'assets')]

# Only include serial pieces you need
hidden = [
    'serial',                     # core pyserial
    'serial.tools.list_ports',
    'serial.urlhandler.protocol_loop',
    'serial.win32', 'serial.serialwin32', 'serial.serialutil',
]
# If you actually need TCP (rfc2217), uncomment the next line:
# hidden += ['serial.urlhandler.protocol_socket','serial.urlhandler.protocol_rfc2217']

excludes = [
    'PySide6.QtWebEngineCore', 'PySide6.QtWebEngineWidgets',
    'PySide6.QtQml', 'PySide6.QtQuick',
    'PySide6.QtMultimedia', 'PySide6.QtPositioning',
    'PySide6.QtCharts', 'PySide6.QtPdf',
]

a = Analysis(
    ['main.py'],
    pathex=[str(HERE)],
    hiddenimports=hidden,
    datas=datas,
    excludes=excludes,
)

pyz = PYZ(a.pure, a.zipped_data, optimize=2)

exe = EXE(
    pyz, a.scripts,
    name='473SerialTerm',
    icon=str(ICO),
    console=False,
    uac_admin=False,
    uac_uiaccess=False,
    version=str(VER),
    upx=False,                 # avoid packer heuristics
)

coll = COLLECT(
    exe, a.binaries, a.zipfiles, a.datas,
    name='473SerialTerm',
    upx=False,                 # avoid packer heuristics (again)
)
