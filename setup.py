'''
	Name: Executable Setup file
	Version: 0.0
	Author: Neal Bazirake
	Description: This file sets up the executable necessary to run the file as an icon on your desktop
'''


from cx_Freeze import setup, Executable
from distutils.core import setup
import py2app

includes = ["re"]

exe = Executable(script="Exmpt_Lbl_Tul.py",base="Win32GUI")
 
setup( name = "Compliance Tool", version = "0.1",
    description = "Tool useful to make report generation of exempt label usage in the Compliance db",
    executables = [exe])
'
