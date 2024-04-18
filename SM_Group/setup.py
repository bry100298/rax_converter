import sys
from cx_Freeze import setup, Executable

# Executable
executables = [Executable("sm_grp.py", base=None)]

# Setup
setup(
    name="SM_GrpConverter",
    version="1.0",
    description="SM Group Converter",
    options={"build_exe": {"packages": ["os", "shutil", "pandas", "xml.etree.ElementTree", "time"], 
                           "include_files": [],
                           "target_name": "SM_GrpConverter.exe"}},  # Specify the target name here
    executables=executables
)


# from distutils.core import setup
# import py2exe

# setup(
#     console=['sm_grp.py']  # Replace 'your_script.py' with the name of your Python script
# )
