# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at https://mozilla.org/MPL/2.0/.

from setuptools import setup, find_packages
from os import listdir

descr = "A declarative API for composing spreadsheets from python"
name = 'xlcompose'
url = 'https://github.com/jbogaardt/xlcompose'
version='0.1.9' # Put this in __init__.py

data_path = ''
setup(
    name=name,
    version=version,
    maintainer='John Bogaardt',
    maintainer_email='jbogaardt@gmail.com',
    packages=['{}.{}'.format(name, p) for p in find_packages(where=name)]+['xlcompose'],
    scripts=[],
    url=url,
    download_url='{}/archive/v{}.tar.gz'.format(url, version),
    license='LICENSE',
    include_package_data=True,
    description=descr,
    # long_description=open('README.md').read(),
    install_requires=[
        "pandas<1.0",
        "xlsxwriter>=1.1.8",
    ],
)
