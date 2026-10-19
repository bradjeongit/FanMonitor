# -*- mode: python ; coding: utf-8 -*-
# 미니 제어기 실행파일 (Windows):  pyinstaller mini_control.spec
# 결과: dist/MINI_CONTROL.exe (파일 하나). 로그는 exe 옆 logs 폴더, 항목 목록은 exe 옆 mini_control.json

a = Analysis(
    ['tools/mini_control.py'],
    pathex=['.', 'tools'],
    binaries=[],
    datas=[],
    hiddenimports=['motor_monitor.modbus', 'ecblue_tester'],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='MINI_CONTROL',
    console=False,
    icon='motor_monitor/assets/motor.ico',
)
