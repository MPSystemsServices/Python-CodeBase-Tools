## Installation and Usage Instructions
# Installation
As of this release, we still do not supply components in the Python Package Index for installation using Python PIP.  But installation should be relatively simple to implement the Python-Codease_Tools on your system.
Follow these steps:
1. Navigate to the site-packages directory either in your primary Python installation or in the appropriate virtual environment.
1. Create a subdirectory named codebasetools
1. Copy all files from this repository's CBToolsInstallDir\codebasetools directory into that new directory
1. In the site-packages directory create a new text file named CodeBaseTools.pth
1. Edit your new .pth file and type in the simple string codebasetools (for an example of what this file should look like, see the CBToolsInstallDir

# Usage in your applications
To create an instance of the Code Base Tools, you simply import the module and create an instance of the module object.  For example:

`from CodeBaseTools import cbTools`

`cbt = cbTools()`

To test that you have a working version of the module, type:

`cbt.getSystemModuleName()`

That should respond with the name of the PYD file being used to access CodeBase(tm).

