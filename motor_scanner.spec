# -*- mode: python ; coding: utf-8 -*-
# 모터 스캐너 실행파일 (Windows):  pyinstaller motor_scanner.spec
# 결과: dist/MOTOR_SCANNER.exe (파일 하나). 로그는 exe 옆 logs 폴더, 마지막 입력값은 motor_scanner.json

a = Analysis(
    ['tools/motor_scanner.py'],
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
    name='MOTOR_SCANNER',
    console=False,
    icon='motor_monitor/assets/motor.ico',
)
