from setuptools import setup, find_packages
import os

setup(
    name='mkreport-create',
    version='0.01',
    packages=find_packages(),
    long_description="This script can be used to generate reports from makefile output and PVS Studio csv reports",


    scripts=[
        os.path.join('mkreport-create', 'mkreport-create.py'),
        os.path.join('mkreport-create', 'summary-diff.py')
    ],

    install_requires=[
        'XlsxWriter',
    ]
)