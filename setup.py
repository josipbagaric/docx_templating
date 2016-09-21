# -*- coding: utf-8 -*-

from setuptools import setup, find_packages


with open('README.rst') as f:
    readme = f.read()

with open('LICENSE') as f:
    license = f.read()

setup(
    name='docx_templating',
    version='0.0.1',
    description='Library for exporting Word files in the desired template.',
    long_description=readme,
    author='Josip Bagaric',
    author_email='bagaricjos@gmail.com',
    url='https://github.com/Bagaric/docx_templating',
    license=license,
    packages=find_packages(exclude=('tests', 'docs'))
)
