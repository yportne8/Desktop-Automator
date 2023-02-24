""" 
(The following are examples for individuals who have placed their 
 python311 folder onto to their %PATH% environment variable. Those
 without python.exe included in PATH will have to use the following:
 start "%LOCALAPPDATA%\Programs\Python311\python.exe setup.py bdist_wheel")

[Making the Process Class into a Installable Package]
 The following are template strings that can be populated via string.format(values).
 This strings populate the below file struture. Following content population, the
 package can be installed as an executable from the command line:

 ```bash
 cd "c:\\path\\to\\Package Name"
 python setup.py bdist_wheel
 ```
 
 This will create two folders in your package directory:
    build and dist. 
 
 '''bash
 cd into the dist folder, then install the package:
 python pip install name-of-your-packge-1.0.0.whl
 '''

 Your package is now installed on your system!
 You can now execute the package as an api through the Console User Interface,
 with all exposed functions as menu options:

 ```python
 python -m name_of_your_package
 ```
 
 At the conclusion of this process you should zip build and dist folders together
 into a version numbered folder and retain the release. This allows for rollbacks
 if bugs arise in future updates. A rollback is as simple as unzipping the 
 versioned folder and installing the file in dist with python pip. Python will
 automatically uninstall the currently installed version and re-install the rollback.
 If a patch needs to be made, your previous version is available in build as reference.

 The file and folder structure of a wrapped Process:
    : \Package Name : Main Folder
    : \Package Name\docs : if available
    : \Package Name\licenses : if necessary, any pypi package can be used as a component for the package
    : \Package Name\requirements.txt : a list of required packages, e.g. process==0.1.0.
    : \Package Name\setup.py : a template file which only requires changes to the package name and specs.
    : \Package Name\pypackage.toml : a template file which only requires changes to the package name.
    : \Package Name\name_of_your_package\__init__.py : requires one line of code: 'from .__main__ import Process'
    : \Package Name\name_of_your_package\__main__.py : contains class Process, using this template.
    : \Package Name\name_of_your_package\other_files_if_necessary_for_readability.py
"""

REQUIREMENTS = ("""
process=={} # must be installed from released whl
""", ["process_version"])


PYPROJECT_TOML = ("""
[build-system]
requires = ["setuptools>=61.0"]
build-backend = "setuptools.build_meta", 

[project]
name = "{}"
version = "{}"
authors = [
  { name="{}", email="{}" },]
description = "{}"
requires-python = "==3.11.1"
classifiers = [
    "Programming Language :: Python :: 3",
    "Operating System :: Microsoft :: Windows"]
""", ["pkg_name", "version", "author", "email", "description"])


SETUP_PY = ("""
from setuptools import setup
from {%pkg_name} import __version__

setup(
    name="{}",
    version=__version__.version,
    author="{}",
    description="{}",
    python_requires="==3.11.1", # being strict here since the Setup.msi bundles 3.11.1
    packages={},
    install_requires={},
    package_data={},
    )
""", ["pkgname", "author", "shortdesc", "packages", "dependencies", "extras"])


INIT_PY = """
from process import Process

# import other_packages_libraries_here
# additional libraries can be installed and imported on the command line:
# 
# ```bash
# c:\\path\\to\\python.exe c:\\path\\to\\pip.exe install pypi_package_name
# ``` 
#
# Be sure to resolve any potential licensing issues prior to installation,
# licenses for installed packages are included in the package repository
# and its contents can simply be copied into a text file and placed into
# the 'license' folder. 
#
# Any external packages along with the package version needs to be added
# to the requirements.txt file of the current directory. This is so that 
# dependency issues are not raised if the Process package and/or python 
# needs to be reinstalled.
#
# [Making Functions Available as Api on the Console] to make Process specific 
# functions availble through the console selection menu, __all__ must be 
# filled in with a string of comma separated function_names. The selection 
# menu will also not list functions that are .__hidden or ._partially_hidden
# even if they are included in __all__.
# 
# If any of the @abstractmethod cannot be filled in for whatever reason,
# you can simply put a place holder method with the same name, and the
# word 'pass' under it. The @abstractmethod does not need to be retained.
#    
# Any number of functions can be added a Process. These functions
# can be made inaccessible outside of the running process, or exposed for
# other operations. As each process is installed separately, processes can 
# import from each other and can even be chained together into a larger 
# processes. Watch out for circular imports when attempting this.
#    
# A login collection process is built into the abstract class. The 
# credentials collected are the user's Global Source Username and 
# Password. User credentials are collected to access the O365.SharePoint
# rest-api, however, authentication happens on a site by site basis, 
# after the assignment of self.sharepoint.site:
#
# self.sharepoint.site = "HR-AMS2"
#
# If access to multiple sites is needed, and are all are accessible
# by the UserID first assigned at startup, then the user does not 
# need to reauthenticate, they can simply reassign .site and access
# the new site's Shared Documentsand as the web context will auto reload. 
# 
# The collection of credentials at startup can be circumvented on the 
# user's device by adding a username and password to the username and
# password parameters at init.


class Process(Process):


    def __init__(self, username: str, password: str):
        super().__init(username, password)

    def dataIn(self):
        try:
            
            # pass can be replaced with code or calls to other functions.
            
            pass
        
        except Exception as e:
            return e

    def dataOut(self, *args, **kwargs):

        try:
            
            # pass can be replaced with code or calls to other functions.
            
            pass
        
        except Exception as e:
            return e
"""


MAIN_PY = """
import traceback
from process.__main__ import Console

# This file should not be changed.


class Console(Console):

    def __init__(self, Process):
        super().__init__(Process)
        assert self.api


if __name__ == "__main__":

    # Imports Process from the inheriting package as a parameter for Console
    from .__init__ import Process 

    try:
        app = Console(Process)
        app.main()
    
    except Exception as e:
        traceback.print_exception(e)
"""