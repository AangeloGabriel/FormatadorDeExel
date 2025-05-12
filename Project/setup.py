from cx_Freeze import setup, Executable
import sys

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="Formatador_Excel",
    version="1.0",
    description="App de formatação de Excel",
    options={
        "build_exe": {
            "packages": ["Functions", "Resources", "ui"],
        }
    },
    executables=[Executable("main.py", base=base)]
)
