# -*- mode: python ; coding: utf-8 -*-
#
# Seoul_Bus_Drive_Recorder v1.19 PyInstaller spec
# - onefile / windowed 빌드 유지 (이전 버전과 동일한 저용량 빌드 방식)
# - 아이콘: Windows는 assets/bus.ico, macOS는 assets/bus.icns 사용
#   (기존 저장소에 없던 Icon.ico 참조 문제 해결)
# - licenses/ 폴더를 --add-data로 배포물에 동봉 (LGPL/GPL/MPL/MIT 전문 + THIRD_PARTY_LICENSES.md)
# - 이미지(아이콘 PNG, 공공누리 마크)는 모두 소스에 base64로 내장되어 있으므로
#   런타임 이미지 자산에 대한 --add-data는 필요 없음.

import sys

app_icon = 'assets/bus.icns' if sys.platform == 'darwin' else 'assets/bus.ico'

a = Analysis(
    ['Seoul_Bus_Drive_Recorder_v1_19.py'],
    pathex=[],
    binaries=[],
    datas=[('licenses', 'licenses')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Seoul_Bus_Drive_Recorder_v1_19',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=[app_icon],
)
