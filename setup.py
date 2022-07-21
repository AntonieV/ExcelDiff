from setuptools import setup

setup(
    name='ExcelDiff',
    version='0.1.0',
    description='A tool to compare two excel files with annotation of the differences.',
    url='https://github.com/AntonieV/ExcelDiff',
    author='Antonie Vietor',
    author_email='a.vietor@gmx.net',
    license='GPLv3',
    packages=['exceldiff'],
    install_requires=[
        'argparse',
        'pathlib',
        'pandas',
        'numpy'
        # 'argparse>=1.4',
        # 'pathlib>=1.0',
        # 'pandas>=1.4',
        # 'numpy>=1.23'
    ],
    classifiers=[
        'Development Status :: 1 - Planning',
        'Programming Language :: Python :: 3',
        'Operating System :: POSIX :: Linux',
        'License :: OSI Approved :: MIT License'
    ],
    entry_points={
        'console_scripts': ['exceldiff=exceldiff.__main__:main'],
    },
)
