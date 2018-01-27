from cx_Freeze import setup, Executable
import os.path


PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')


include = ['win32com', 'xlwings','numpy']
exclude = ['scipy', 'xml', 'Tkinter', 'Tkconstants', 'pydoc', 'tcl',
          'tk', 'matplotlib', 'PIL', 'nose', 'setuptools', 'xlrd', 'xlwt', 'PyQt4', 'markdown',
          'IPython', 'docutils','mkl']

buildOptions = dict(packages = include, #excludes = exclude, 
                    optimize = 2,
                    include_files=[
                            os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll'),
                            os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll')
                            ],
                    excludes = exclude)

base = "Win32GUI"

executables = [
        Executable('import_equipment.py', base=base),
        Executable('indicators_to_checklist.py', base=base)        
               ]

setup(name='import_equipment',
      version = '0.1',
      options = dict(build_exe = buildOptions),
      executables = executables)
