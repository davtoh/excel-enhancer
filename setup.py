import sys
import io
import os
import re
from setuptools import setup, find_packages


def read(*names, **kwargs):
    with io.open(
        os.path.join(os.path.dirname(__file__), *names),
        encoding=kwargs.get("encoding", "utf8")
    ) as fp:
        return fp.read()


def find_version(*file_paths):
    version_file = read(*file_paths)
    version_match = re.search(r"^__version__ = ['\"]([^'\"]*)['\"]",
                              version_file, re.M)
    if version_match:
        return version_match.group(1)
    raise RuntimeError("Unable to find version string.")


version = find_version('excel_enhancer.py')
here = os.path.abspath(os.path.dirname(__file__))

long_description = read(os.path.join(here, 'README.md'))

setup(
    name="excel-enhancer",
    description="Enhance Excel with Python User Defined Functions (UDFs) and pandas",
    version=version,
    author="David Toro",
    author_email="davsamirtor@gmail.com",
    url="https://github.com/davtoh/excel-enhancer",
    packages=find_packages(exclude=['tests*']),
    py_modules=["excel_enhancer", "core", "mod_udf"],
    license="BSD-3-Clause",
    long_description=long_description,
    long_description_content_type="text/markdown",
    python_requires=">=3.9",
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'Intended Audience :: Financial and Insurance Industry',
        'Intended Audience :: Science/Research',
        'Intended Audience :: End Users/Desktop',
        'License :: OSI Approved :: BSD License',
        'Operating System :: Microsoft :: Windows',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
        'Programming Language :: Python :: 3.12',
        'Programming Language :: Python :: 3.13',
        'Programming Language :: Python :: 3.14',
        'Topic :: Office/Business :: Financial :: Spreadsheet',
        'Topic :: Software Development :: Libraries :: Python Modules',
    ],
    keywords='excel, udfs, xlwings, pandas, data-analysis',
    platforms=['win32', 'win_amd64'],
    install_requires=[
        'numpy>=1.24.0',
        'pandas>=2.0.0',
        'xlwings>=0.30.0',
        'scipy>=1.10.0',
        'pywin32>=300; platform_system=="Windows"',
    ],
    extras_require={
        'dev': ['pytest>=8.0.0', 'faker>=18.0.0'],
        'extensions': ['faker>=18.0.0'],
    },
    entry_points={
        'console_scripts': [
            'excel-enhancer = excel_enhancer:serve',
        ],
    }
)
