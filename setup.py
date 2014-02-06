#!/usr/bin/env python

try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup
from glob import glob

# Make data go into site-packages (http://tinyurl.com/site-pkg)
from distutils.command.install import INSTALL_SCHEMES
for scheme in INSTALL_SCHEMES.values():
    scheme['data'] = scheme['purelib']

DESCRIPTION = (
    'The docx module creates, reads and writes Microsoft Office Word 2007 do'
    'cx files'
)

setup(
    name='docx',
    version='0.2.4',
    install_requires=['lxml', 'Pillow>=2.0'],
    description=DESCRIPTION,
    author='Mike MacCana',
    author_email='python-docx@googlegroups.com',
    maintainer='Steve Canny',
    maintainer_email='python-docx@googlegroups.com',
    url='http://github.com/mikemaccana/python-docx',
    py_modules=['docx'],
    data_files=[
        ('docx-template/_rels',      glob('template/_rels/.*')),
        ('docx-template/docProps',   glob('template/docProps/*.*')),
        ('docx-template/word',       glob('template/word/*.xml')),
        ('docx-template/word/theme', glob('template/word/theme/*.*')),
    ],
)
