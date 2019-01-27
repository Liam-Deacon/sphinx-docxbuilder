#!/usr/bin/env python
# -*- coding: utf-8 -*-

import io
import os
import re
import datetime

try:
    from setuptools import find_packages, setup
except ImportError:
    from distutils.core import setup

    def find_packages(path='.', **kwargs):
        ret = []
        for root, dirs, files in os.walk(path):
            if '__init__.py' in files:
                ret.append(re.sub('^[^A-z0-9_]+', '', root.replace('/', '.')))
        return ret

# Package meta-data.
DESCRIPTION = ('An extension for a docx file generation with Sphinx '
               '(Fork of https://bitbucket.org/haraisao/sphinx-docxbuilder)')

# Import the README and use it as the long-description.
# Note: this will only work if 'README.md' is present in your MANIFEST.in file!
with io.open(os.path.join('README'), encoding='utf-8') as f:
    long_description = '\n' + f.read()

setup(
    name='sphinx-docxbuilder',
    version=datetime.date.today().strftime(r'%Y.%m.%d'),
    description=DESCRIPTION,
    long_description=long_description,
    long_description_content_type='text/markdown',
    author='Isao Hara',
    author_email='isao-hara@aist.go.jp',
    maintainer='Liam Deacon',
    maintainer_email='liam.m.deacon@gmail.com',
    python_requires='Python>=2.7',
    url='https://github.com/Lightslayer/sphinx-docxbuilder',
    packages=find_packages(exclude=['tests']),
    license='MIT License',
    classifiers=[
        # Trove classifiers
        # Full list: https://pypi.python.org/pypi?%3Aaction=list_classifiers
        'Programming Language :: Python',
        'Framework :: Sphinx :: Extension',
        'Topic :: Documentation :: Sphinx'
    ]
)
