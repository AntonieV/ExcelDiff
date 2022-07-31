"""A tool to compare two excel files with annotation of the differences."""
import os
from setuptools import setup
# requirements from requirements.txt
REQUIREMENTS = [i.strip() for i in open("requirements.txt").readlines()]


# Utility function to read the README file.
# see https://pythonhosted.org/an_example_pypi_project/setuptools.html
# Used for the long_description.  It's nice, because now 1) we have a top level
# README file and 2) it's easier to type in the README file than to put a raw
# string in below ...
def read(fname):
    """Linking README to package setup"""
    return open(os.path.join(os.path.dirname(__file__), fname)).read()


setup(
    name='SpreadSheetDiff',
    version='0.1.0',
    author='Antonie Vietor',
    author_email='a.vietor@gmx.net',
    description='A tool to compare two excel files with annotation of the differences.',
    url='https://github.com/AntonieV/SpreadSheetDiff',
    license='GPLv3',
    packages=['spreadsheetdiff'],
    python_requires='>3.6',
    install_requires=REQUIREMENTS,
    long_description=read('README.md'),
    classifiers=[
        'Development Status :: 1 - Planning',
        'Programming Language :: Python :: 3',
        'Operating System :: POSIX :: Linux',
        'License :: OSI Approved :: MIT License'
    ],
    entry_points={
        'console_scripts': ['spreadsheetdiff=spreadsheetdiff.main:main'],
    },
)
