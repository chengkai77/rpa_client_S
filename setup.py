from distutils.core import setup
from Cython.Build import cythonize
import os

file_path = os.getenv("file_path")
try:
    setup(
        # name='Hello world app',
        ext_modules=cythonize(file_path),
        # build_ext="build_ext"
    )
except Exception as e:
    # print(e)
    pass
