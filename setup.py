try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

import os.path
import sys

PY3 = sys.version_info >= (3, 0)

setup(
    name='gspread-formatting',
    packages=['gspread_formatting'],
    package_data={'': ['*.rst']},
    test_suite='test',
    install_requires=[
        'gspread>=3.0.0' 
        ],
    description='Complete Google Sheets formatting support for gspread worksheets',
    author='Robin Thomas',
    author_email='rthomas900@gmail.com',
    license='MIT',
    url='https://github.com/robin900/gspread-formatting',
    keywords=['spreadsheets', 'google-spreadsheets', 'formatting', 'cell-format'],
    classifiers=[
        "Programming Language :: Python",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: Implementation :: CPython",
        "Programming Language :: Python :: Implementation :: PyPy",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers",
        "Intended Audience :: Science/Research",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
        "Topic :: Software Development :: Libraries :: Python Modules"
        ],
    zip_safe=True
)
