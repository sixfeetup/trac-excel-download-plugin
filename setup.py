#!/usr/bin/env python
# -*- coding: utf-8 -*-

from setuptools import setup, find_packages

extra = {}
try:
    import babel
    extra['message_extractors'] = {
        'tracexceldownload': [
            ('**/*.py',              'python', None),
            ('**/templates/**.html', 'genshi', None),
        ],
    }
    from trac.util.dist import get_l10n_cmdclass
    extra['cmdclass'] = get_l10n_cmdclass()
except ImportError:
    pass

setup(
    name = 'TracExcelDownload',
    version = '0.12.0.8',
    description = 'Allow to download query and report page as Excel',
    license = 'BSD', # the same as Trac
    packages = find_packages(exclude=['*.tests*']),
    package_data = {
        'tracexceldownload': [
            'locale/*.*', 'locale/*/LC_MESSAGES/*.mo',
        ],
    },
    test_suite = 'tracexceldownload.tests.suite',
    install_requires = ['Trac'],
    extras_require = {'openpyxl': 'openpyxl', 'xlwt': 'xlwt'},
    entry_points = {
        'trac.plugins': [
            'tracexceldownload.api = tracexceldownload.api',
            'tracexceldownload.ticket = tracexceldownload.ticket',
            'tracexceldownload.translation = tracexceldownload.translation',
        ],
    },
    **extra)
