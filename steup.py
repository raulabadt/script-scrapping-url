import sys
from cx_Freeze import setup, Executable

# Reemplaza 'tu_script.py' con el nombre de tu script Python
script = 'scriptentorno.py'

base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

executables = [Executable(script, base=base)]

setup(
    name='NombreDelEjecutable',
    version='1.0',
    description='Descripción del ejecutable',
    executables=executables
)
