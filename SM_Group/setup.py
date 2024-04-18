from cx_Freeze import setup, Executable

setup(
    name="SM_GrpConverter",
    version="1.0",
    description="SM Group",
    executables=[Executable("sm_grp.py")],
)

# from distutils.core import setup
# import py2exe

# setup(
#     console=['sm_grp.py']  # Replace 'your_script.py' with the name of your Python script
# )
