from process import __version__
from setuptools import setup


setup(
    name='dxc',
    version=__version__.version,
    author='Chong Kim',
    author_email='ckim32@dxc.com',
    description='A collection of easy to use programmatic tools to assist in the execution of existing desktop scripts.',
    long_description='Provides programmtic access to the O365-rest-api, automation of desktop apps, along with prebuilt tools that that can be easily attached to any existing VBA script.',
    python_requires='==3.11.1', # being strict here since the Setup.msi bundles 3.11.1
    packages=['dxc'],
    install_requires=[
        'xlrd==2.0.1',
        'sphinx==5.0.2',
        'pandas==1.5.2',
        'pyodbc==4.0.35',
        'openpyxl==3.0.10',
        'msoffcrypto-tool==5.0.0',
        'office365-rest-python-client==2.3.16',]
    )