from cx_Freeze import setup, Executable
from pip.operations import freeze
import os

os.environ['TCL_LIBRARY'] = r'C:\Users\grlurton\AppData\Local\Continuum\anaconda3\tcl\tcl8.6'
os.environ['TK_LIBRARY'] = r'C:\Users\grlurton\AppData\Local\Continuum\anaconda3\tcl\tk8.6'

include = ['pandas', 'numpy', 'dateutil', 'six', 'pytz', 'xlwings', 'comtypes', 'tkinter']

#exclude = ['babel', 'bokeh', ]

exclude = ['backports', 'asyncio' , 'bs4', 'concurrent', 'ctypes', 'curses', 'distutils', 'email','html','PIL','multiprocessing','IPython','json','pkg_resources','pydoc_data','sqlite3','tcl.tzdata','xml','urllib','sklearn','http','test',
            'unittest', 'win32com']
x = freeze.freeze()

for p in x :
    package = p.split('==')
    package = package[0]
    if package not in include:
        exclude.append(package)
        exclude.append(package.lower())

buildOptions = dict(packages = include, excludes = exclude, optimize = 2)

base = 'Console'

executables = [
    Executable('simplified_data_collection.py', base=base)
]

setup(name='ltqi_data_collection',
      version = '0.1',
      description = 'bar',
      options = dict(build_exe = buildOptions),
      executables = executables)
