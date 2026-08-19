# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# spec 안의 상대경로는 spec 파일 위치를 기준으로 잡힌다.
# 이 spec 은 '프로그램 만들기' 안에 있으므로 한 단계 위가 프로젝트 폴더다.
ROOT = os.path.dirname(SPECPATH)

block_cipher = None

_hidden = []
_datas  = []

for _pkg in [
    'customtkinter',      # 테마 json · 글꼴 (없으면 exe 가 안 켜진다)
    'tkinterdnd2',        # tkdnd 바이너리 (없으면 드래그앤드롭이 죽는다)
    'googleapiclient',
    'google.auth',
    'google.oauth2',
    'google_auth_oauthlib',
    'google.auth.transport',
    'httplib2',
    'uritemplate',
]:
    try:
        _hidden += collect_submodules(_pkg)
        _datas  += collect_data_files(_pkg)
    except Exception:
        pass

a = Analysis(
    [os.path.join(ROOT, 'run.py')],              # 실행 파일
    # 앱 코드(core/engine/theme/screens/app/updater)는 일부러 넣지 않는다.
    # exe 옆 '프로그램' 폴더에서 읽어야 원격 갱신이 적용된다.
    pathex=[],

    binaries=[],
    datas=_datas,
    hiddenimports=_hidden + [
        'google',
        'google.auth',
        'google.auth.credentials',
        'google.auth.transport',
        'google.auth.transport.requests',
        'google.oauth2',
        'google.oauth2.credentials',
        'google_auth_oauthlib',
        'google_auth_oauthlib.flow',
        'googleapiclient',
        'googleapiclient.discovery',
        'googleapiclient.http',
        'googleapiclient.errors',
        'httplib2',
        'uritemplate',
        'customtkinter',
        'tkinterdnd2',
        'customtkinter.windows',
        'customtkinter.windows.widgets',
        'customtkinter.windows.widgets.core_rendering',
        'customtkinter.windows.widgets.font',
        'customtkinter.windows.widgets.image',
        'customtkinter.windows.widgets.scaling',
        'customtkinter.windows.widgets.theme',
        'customtkinter.windows.widgets.utility',
                'win32com',
        'win32com.client',
        'pythoncom',
        'win32com.server',
        'pywintypes',
        'win32api',
        'openpyxl',
        'openpyxl.cell',
        'openpyxl.styles',
                'email',
        'email.mime',
        'email.mime.multipart',
        'email.mime.text',
        'email.mime.base',
        'email.encoders',
        'email.message',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='메일자동화',
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
)
