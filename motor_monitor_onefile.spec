# -*- mode: python ; coding: utf-8 -*-
# 운영 프로그램 실행파일 - 파일 하나 (GitHub Actions 가 Release 'exe-latest' 에 올리고 update.bat 이 내려받음)
#   pyinstaller motor_monitor_onefile.spec  ->  dist/BLDC_MOTOR_MONITOR.exe
# Config.dll 과 logs 폴더는 exe 가 있는 폴더를 씀 (프로그램 폴더 맨 위에 두면 python run.py 와 같은 Config.dll 사용)
# 폴더형 빌드(exe + 여러 파일)는 motor_monitor.spec / build_exe.bat

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
    a.binaries,
    a.datas,
    [],
    name='BLDC_MOTOR_MONITOR',
    console=False,
    icon='motor_monitor/assets/motor.ico',
)
