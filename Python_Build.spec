# -*- mode: python ; coding: utf-8 -*-
#
# Seoul_Bus_Drive_Recorder PyInstaller spec
# - onefile / windowed 빌드 유지 (이전 버전과 동일한 저용량 빌드 방식)
# - 아이콘: Windows는 assets/bus.ico, macOS는 assets/bus.icns 사용
#   (기존 저장소에 없던 Icon.ico 참조 문제 해결)
# - licenses/ 폴더를 --add-data로 배포물에 동봉 (LGPL/GPL/MPL/MIT 전문 + THIRD_PARTY_LICENSES.txt)
# - 이미지(아이콘 PNG, 공공누리 마크)는 모두 소스에 base64로 내장되어 있으므로
#   런타임 이미지 자산에 대한 --add-data는 필요 없음.
# - macOS에서는 단일 바이너리에 아이콘이 적용되지 않으므로,
#   BUNDLE 블록으로 .app 번들을 함께 생성해 bus.icns를 적용한다.

APP_VERSION = "1_22"  # 앱/파일명에 쓰이는 버전 표기 (예: v1_21)

VERSION_DOT = APP_VERSION.replace("_", ".")

MAC_BUNDLE_VERSION = VERSION_DOT

BUNDLE_ID = "com.metapbl.seoulbusdriverecorder"

import sys

APP_NAME = f"Seoul_Bus_Drive_Recorder_v{APP_VERSION}"
SOURCE_SCRIPT = f"{APP_NAME}.py"
DISPLAY_NAME = f"Seoul Bus Drive Recorder v{VERSION_DOT}"

app_icon = 'assets/bus.icns' if sys.platform == 'darwin' else 'assets/bus.ico'

a = Analysis(
    [SOURCE_SCRIPT],
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
    name=APP_NAME,
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

if sys.platform == 'darwin':
    app = BUNDLE(
        exe,
        name=f"{APP_NAME}.app",
        icon=app_icon,
        bundle_identifier=BUNDLE_ID,
        version=MAC_BUNDLE_VERSION,
        info_plist={
            'NSPrincipalClass': 'NSApplication',
            'NSHighResolutionCapable': True,
            'CFBundleDisplayName': DISPLAY_NAME,
            'LSMinimumSystemVersion': '11.0',
        },
    )
