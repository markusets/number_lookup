from distutils.core import setup
import py2exe

setup(
    windows=['area_code_analyzer.py'],
    options={
        "py2exe": {
            "includes": ["tkinter"],
            "icon_resources": [(1, "belmont_crest.ico")]
        }
    }
)
