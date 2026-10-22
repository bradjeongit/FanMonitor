# -*- mode: python ; coding: utf-8 -*-
# 모터 시험기(SUNINUS MOTOR TEST) 실행파일 (Windows):  pyinstaller motor_tester.spec
# 결과: dist/MOTOR_TESTER.exe (파일 하나). 통신 로그는 exe 옆 log 폴더, 설정은 exe 옆 Motor_Config.ini

a = Analysis(
    ['tools/motor_tester.py'],
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
    name='MOTOR_TESTER',
    console=False,
    icon='motor_monitor/assets/motor.ico',
)
