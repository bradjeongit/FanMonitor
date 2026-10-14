# -*- mode: python ; coding: utf-8 -*-
# 기존모터 시험기 실행파일 (Windows):  pyinstaller ecblue_tester.spec
# 결과: dist/ECBLUE_TESTER.exe (파일 하나). 로그는 exe 옆 logs 폴더에 생성

a = Analysis(
    ['tools/ecblue_tester.py'],
    pathex=['.'],
    binaries=[],
    datas=[],
    hiddenimports=['motor_monitor.modbus'],
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
    name='ECBLUE_TESTER',
    console=False,
    icon='motor_monitor/assets/motor.ico',
)
