# Installation and Usage Instructions
## Installation
As of this release, we still do not supply components in the Python Package Index for installation using Python PIP.  But installation should be relatively simple to implement the Python-CodeBase_Tools on your system.
Follow these steps (directory and file names are case sensitive, even on Windows):
1. Navigate to the site-packages directory either in your target Python installation or in the appropriate virtual environment.
1. Create a subdirectory within site-packages named codebasetools
1. Copy all files from this repository's CBToolsInstallDir\codebasetools directory into that new directory
1. In the site-packages directory create a new text file named CodeBaseTools.pth
1. Edit your new .pth file and type in the simple string codebasetools (for an example of what this file should look like, see the CBToolsInstallDir

## Usage in your applications
To create an instance of the Code Base Tools component, you simply import the module and create an instance of the module object.  For example:

`from CodeBaseTools import cbTools`

`cbt = cbTools()`

To test that you have a working version of the module, type:

`cbt.getSystemModuleName()`

That should respond with the name of the PYD file being used to access CodeBase(tm).

Extensive documentation is built into the system .py modules, which you can access by the pydoc Python module.  Alternatively there is a suite of HTML help files that provide detailed instructions for all system functions and features.  See the [`MPSSCommon Module Docs.html`] ("Documentation\MPSSCommon Module Docs.html") file in the Documentation subdirectory.

