import sys
from cx_Freeze import setup, Executable

build_exe_options = {'include_files': ['Hawk.png', 'Hawk.ico', 'Hawk-doc.pdf', 'hcad-hawk3.ini', ('hcad.py', 'lib/hcad.py'), 'hcad.xlsx']}

# GUI applications require a different base on Windows (the default is for a console application).
base = None

if sys.platform == "win32":
    base = "Win32GUI"

setup(  name = "Hawk",
        version = "3.0",
        description = "Ultimate Site Grabber",
		options = {'build_exe': build_exe_options},
        executables = [Executable(script="Hawk3.py", base=base, icon="Hawk.ico")])