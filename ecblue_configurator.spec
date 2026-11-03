# -*- mode: python ; coding: utf-8 -*-
# ECblue configurator 실행파일 (Windows):  pyinstaller ecblue_configurator.spec
# 결과: dist/ECBLUE_CONFIGURATOR.exe (파일 하나). 로그는 exe 옆 logs 폴더, 최근 IP·ID 는 exe 옆 ecblue_configurator.json
# 레지스터 정의·통신은 모터 스캐너(tools/motor_scanner.py)를 함께 묶음

a = Analysis(
    ['tools/ecblue_configurator.py'],
    pathex=['.', 'tools'],
    binaries=[],
    datas=[],
    hiddenimports=['motor_monitor.modbus', 'motor_scanner'],
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
    name='ECBLUE_CONFIGURATOR',
    console=False,
    icon='motor_monitor/assets/motor.ico',
)
