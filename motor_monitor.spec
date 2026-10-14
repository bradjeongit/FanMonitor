# -*- mode: python ; coding: utf-8 -*-
# 실행파일 만들기 (Windows):  pyinstaller motor_monitor.spec
# 결과: dist/BLDC_MOTOR_MONITOR/BLDC_MOTOR_MONITOR.exe  (Config.dll 은 exe 옆에 두고, logs 폴더는 exe 옆에 생성)

a = Analysis(
    ['run.py'],
    pathex=[],
    binaries=[],
    datas=[('motor_monitor/assets', 'motor_monitor/assets')],
    hiddenimports=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='BLDC_MOTOR_MONITOR',
    console=False,
    icon='motor_monitor/assets/motor.ico',
)
coll = COLLECT(exe, a.binaries, a.datas, name='BLDC_MOTOR_MONITOR')
