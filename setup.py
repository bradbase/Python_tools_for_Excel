"""
Python_tools_for_Excel

"""

import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="Python_tools_for_Excel",
    version="0.0.1b0",
    author="Bradley van Ree",
    author_email="flyingkoala@bradbase.net",
    description="Python_tools_for_Excel",
    long_description=long_description,
    long_description_content_type="text/markdown",
    keywords=['xls',
        'excel',
        'spreadsheet',
        'workbook',
        'vba',
        'macro',
        'data analysis',
        'analysis'
        'reading excel',
        'excel formula',
        'excel formulas',
        'excel equations',
        'excel equation',
        'formula',
        'formulas',
        'equation',
        'equations',
        'pandas',
        'harvest',
        'timeseries',
        'time series',
        'energy',
        'accounting',
        'research',
        'visualization',
        'scenario analysis',
        'modelling',
        'model',
        'unit testing',
        'testing',
        'audit'],
    url="https://github.com/bradbase/Python_tools_for_Excel",
    packages=setuptools.find_packages(),
    classifiers=[
        "License :: OSI Approved :: 3-Clause BSD",
        "Operating System :: OS Independent",
    ],
    install_requires=['msmb_theme']
)
