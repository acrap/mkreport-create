from setuptools import setup, find_packages
import os

setup(
    name='mkreport_create',
    version='0.3',
    packages=find_packages(),
    long_description="The Python3 script can be used to make reports from makefile output and PVS Studio csv reports",
    author = "Andrey Strunin",
    license="MIT",

    scripts=[
        os.path.join('bin', 'mkreport_create.py'),
        os.path.join('bin', 'mksummary_diff.py')
    ],

    install_requires=[
        'XlsxWriter',
    ]
)

