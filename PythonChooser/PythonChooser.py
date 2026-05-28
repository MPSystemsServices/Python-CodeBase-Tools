"""
Utility to activate one of multiple version instances of Python on a computer by setting the system PATH 
Environment Variable to point to the working directories for that instance.  This works at a level above
the venv system.  It assumes that each version instance has its own system wide site-packages directory
in the lib directory under the root directory for the specific version of Python.  It looks for directories
on C: and D: that contain python.exe immediately below the drive root directory.  For example it will find
python.exe in c:\\python310 or d:\\python_32_310 or c:\\python310_64, etc.

It than presents the user a list of available Python installations with the corresponding sub versions
implemented and allows the selection of one (or none) of them.  It then removes all references to Python
pathing in the PATH Environment Variable, and inserts new path elements for the selected instance of 
Python.  It makes reference to a PythonChooser.ini file for a list of directories into which it needs to
write a specific version of pyvenv.cfg to tell other applications where to look for their Python instance.
"""
from easygui import passwordbox

try:
    import stat
    import sys
    import os
    import glob
    import shutil
    import ctypes
    import datetime
    import win32api
    import win32con
    import easygui
    import time

except ImportError:
    print("Missing library in current Python instance.  Needs easygui and win32api (pywin32) from PPI.")
    exit()
cVersionTest = sys.version
cVersionStr = str(sys.version_info.major) + "." + str(sys.version_info.minor)
cCodeBaseDir = ""
bIs64Bit = "64 bit" in cVersionTest
if bIs64Bit:
    cVersionStr = cVersionStr + " 64-bit"
if sys.version_info[0] <= 2 or (sys.version_info[0] == 3 and sys.version_info[1] <= 6):
    _ver3x = False
    # import subprocess32 as subprocess
    print("SORRY!  This component will NOT WORK with Python 2.7 or earlier or 3.6 or earlier.  Terminating!")
    exit()
    # Note the problem is the subprocess32 compatibility module which may or may not be present and which
    # does NOT provide full compatibility with version 3.6 and later subprocess.
else:
    _ver3x = True
    import subprocess
cIniInfo = ""
print("VERSION TEST:", cVersionTest)

def ChoosePython(cSelectedPython=""):
    """
    Does all the work
    :return: True
    """
    global cIniInfo
    global cVersionStr
    global cCodeBaseDir
    bOK = True
    nTest = ctypes.windll.shell32.IsUserAnAdmin()
    if nTest != 1:
        bOK = False
        easygui.msgbox("This Python script must be run in Administrator mode.  Terminating.")
        return
    cIniInfo = FILETOSTR("pythonchooser.ini")
    if not cIniInfo:
        bOK = False
        easygui.msgbox("No ini file found. Terminating.")
        return
    if isinstance(cIniInfo, bytes):
        cIniInfo = cIniInfo.decode("utf-8")

    xItem = list()
    xPyVersions = list()
    xTest = ADIR("c:\\python*\\python.exe")
    for xFile in xTest:
        xPyVersions.append([xFile[0], xFile[1], ""])
    xTest2 = ADIR("d:\\python*\\python.exe")
    for xFile in xTest2:
        xPyVersions.append([xFile[0], xFile[1], ""])
    xVersion = list()
    for xItem in xPyVersions:
        xResult = subprocess.run(xItem[0] + " -V -V", capture_output=True)
        if xResult.returncode != 0:
            continue
        cVersionText = xResult.stdout.decode("utf-8")
        if not cVersionText:
            # Probably a 2.x Python version
            xResult = subprocess.run(xItem[0] + " -V", capture_output=True)
            cVersionText = xResult.stderr.decode("utf-8")
            if not cVersionText:
                continue
        xParts = cVersionText.split(" ")
        cBaseVer = xParts[1]
        if "32 bit" in cVersionText:
            cBitness = "32-bit"
        elif "64 bit" in cVersionText:
            cBitness = "64-bit"
        else:
            cBitness = "Unspecified bits"
        xVersion.append(xItem[0] + " - " + cBaseVer + " - " + cBitness)

    cFunctionSelect = ""
    bOK = True
    if len(xItem) == 0:
        bOK = False
    if bOK:
        cMessage = "Select version of Python to make new CURRENT.  \nCurrently version: " + cVersionStr
        cFunctionSelect = easygui.choicebox(msg=cMessage,
                                            title="Python Version Selector", choices=xVersion)
    if cFunctionSelect:
        cCurrentDir = os.getcwd()
        # cCodeBaseDir = easygui.diropenbox(title="CodeBaseTools Source Directory Selector", default="")
        # ABOVE CREATES A PATH PROBLEM THAT UPSETS THE subprocess.call() method elsewhere.
        cCodeBaseDir = "E:\\GIT_Repositories\\Python-CodeBase-Tools\\CBToolsInstallDir\\codebasetools"
        if cCodeBaseDir:
            os.chdir(cCurrentDir)
            SetPythonAsMaster(cFunctionSelect)
    else:
        print("Nothing selected")
        bOK = False
    return bOK


def SetPythonAsMaster(cVersion):
    """
    Does the work of altering the environment variables and paths to support one and only one version of Python.
    Other versions of Python will be ignored by both COM servers and Python scripts run by Python.exe and Pythonw.exe.
    :param cVersion:
    :return: True on OK, else False and displays error message.
    """
    bOK = True
    if not cVersion:
        bOK = False
        print("NO Version of Python Selected. TERMINATING!")
        sys.exit(1)

    xPyParts = cVersion.split(" - ")
    cBasePython = xPyParts[0]
    cBaseVersion = xPyParts[1]
    cBaseBitness = xPyParts[2]
    xBaseParts = os.path.split(cBasePython)
    cBaseDirectory = xBaseParts[0] + "\\"
    print(xBaseParts)
    print("Selected Python Instance", cBasePython)
    print("Selected Python Version", cBaseVersion)
    print("Selected Python Bitness", cBaseBitness)
    print("Selected Directory:", cBaseDirectory)
    xNewPath = list()
    xNewPath.append(cBaseDirectory)
    cScriptPath = os.path.join(cBaseDirectory, "Scripts")
    xNewPath.append(cScriptPath)
    cSitePackages = os.path.join(cBaseDirectory, "Lib", "site-packages")
    xNewPath.append(cSitePackages)
    print(xNewPath)
    ClearPythonPaths()
    SetPythonPaths(xNewPath)
    SetPythonEnv(cBaseDirectory, cBaseVersion, cScriptPath)
    InstallCodeBaseTools(cBaseDirectory, cBaseVersion, cSitePackages)
    bComplete = easygui.ynbox("Do you want to update Python CodeBaseTools?", "Py Chooser")
    if bComplete:
        print("Updating Python CodeBaseTools")
        PopulateCodeBaseToolsFromMaster(cBaseDirectory, cBaseVersion, cSitePackages)
    else:
        print("NO CodeBaseTools update!")
    print ("DONE")
    return bOK

def PopulateCodeBaseToolsFromMaster(cBaseDirectory, cBaseVersion, cSitePackages):
    """
    This copies all the required files from the repo master codebasetools directory into the appropriate
    directory in the site-packages\\codebasetools directory for this version of Python.
    :param cBaseDirectory:
    :param cBaseVersion:
    :param cSitePackages:
    :return:
    """
    global cCodeBaseDir
    cCodeBaseSourceDir = cCodeBaseDir
    cCodeBaseTargetDir = os.path.join(cSitePackages, "codebasetools")
    cSourceSkel = os.path.join(cCodeBaseSourceDir, "*.*")
    xSourceFiles = ADIR(cSourceSkel)
    for xFile in xSourceFiles:
        cFileName = xFile[0]
        cDir, cBase = os.path.split(cFileName)
        cTargetFile = os.path.join(cCodeBaseTargetDir, cBase)
        if os.path.exists(cTargetFile):
            os.remove(cTargetFile)
        shutil.copyfile(cFileName, cTargetFile)
        print("COPIED:", cTargetFile)
    return True

def InstallCodeBaseTools(cBaseDirectory, cBaseVersion, cSitePackages):
    """
    Actually copies the appropriate set of CodeBaseTools files from the development directory(ies) to the
    codebasetools directory in site-packages.  Also sets up the codebasetools.pth file to link into the Python
    module search path.
    :param cBaseDirectory:
    :param cBaseVersion:
    :return: True on OK, else False and displays error message.
    """

    bOK = True
    cPathFile = os.path.join(cSitePackages, "CodeBaseTools.pth")
    try:
        os.remove(cPathFile)
    except FileNotFoundError:
        pass
    except:
        bOK = False
        print("NO CodeBaseTools update! Unidentified Error")

    if not os.path.isfile(cPathFile):
        bOK = STRTOFILE("codebasetools", cPathFile)

    if bOK:
        if os.path.exists(cSitePackages):
            cCBToolsDir = os.path.join(cSitePackages, "codebasetools")
            if not os.path.exists(cCBToolsDir):
                print("MAKING:", cCBToolsDir)
                os.mkdir(cCBToolsDir)
    return bOK

def SetPythonEnv(cBaseDirectory, cBaseVersion, cScriptPath):
    """
    puts the pyvenv.cfg files where they belong with the proper Python version reference.
    :param cBaseDirectory:
    :return: True on OK, else False and displays error message.
    """
    global cIniInfo
    xVenvContent = list()
    xVenvContent.append("home = " + cBaseDirectory)
    xVenvContent.append("include-system-site-packages = true")
    xVenvContent.append("version = " + cBaseVersion)
    xVenvContent.append("executable = " + os.path.join(cBaseDirectory, "python.exe"))
    cVenvContent = "\n".join(xVenvContent)
    xFileList = cIniInfo.split("\n")
    xFileList.append(cBaseDirectory)
    xFileList.append(cScriptPath)
    print("THE INI DIRECTORY LIST:", xFileList)
    for cFile in xFileList:
        cFile = cFile.strip()
        cTargetFile = os.path.join(cFile, "pyvenv.cfg")
        STRTOFILE(cVenvContent, cTargetFile)
        print("Creating PYVENV file:", cTargetFile)
    return True

def ClearPythonPaths():
    """
    reads the environment looking at the PATH value(s) and removes all path elements that have the
    word "python" in them.
    :return: True on OK, else False and displays error message.
    """
    xDelPaths = list()
    cRawPath = os.environ.get("PATH", "NONE")
    if cRawPath != "NONE":
        xPathElems = cRawPath.split(";")
        for cPathElem in xPathElems:
            if "PYTHON" in cPathElem.upper():
                xDelPaths.append(cPathElem)

    if len(xDelPaths) > 0:
        # print("CLEARING IN CURRENT DIRECTORY", os.getcwd())
        cPathManPath = os.getcwd()
        cPathManPath = os.path.join(cPathManPath, "pathman.exe")
        for cDelPath in xDelPaths:
            cWorkProc = cPathManPath + " /rs " + cDelPath
            # print("CLEARING WORK PROC:", cWorkProc)
            subprocess.run(cWorkProc, timeout=None, check=False)
            print("PYTHON PATH REMOVED:", cDelPath)
        bRet = True
    else:
        bRet = False
    return bRet

def SetPythonPaths(xpNewPath):
    """
    uses pathman.exe to define the new paths required for this version of python.
    :param xpNewPath:
    :return:
    """
    if len(xpNewPath) == 0:
        raise ValueError("No Paths to Set!")
    else:
        # print("SETTING IN CURRENT DIRECTORY", os.getcwd())
        cPathManPath = os.getcwd()
        cPathManPath = os.path.join(cPathManPath, "pathman.exe")
        for cPathElem in xpNewPath:
            cWorkProc = cPathManPath + " /as " + cPathElem
            # print("SETTING WORK PROC:", cWorkProc)
            subprocess.run(cWorkProc)
            print("SET NEW PATH:", cPathElem)
    cBasePythonPath = xpNewPath[0]
    cBasePythonPath = cBasePythonPath.strip("\\")
    cWorkProc = "setx PYTHONPATH " + cBasePythonPath + " /M"
    subprocess.run(cWorkProc)
    print("SET NEW PYTHONPATH: ", cBasePythonPath)
    return True

def ADIR(cSkel=""):
    """
    Provides capability like VFP ADIR() function that returns an array with detailed information about every
    file where the directory and name matches the contents of the cSkel parameter -- recognizing * and ? as
    wild card characters in the name.
    :param cSkel: file name to match with full path name (NOT relative paths).  e.g. c:\\temp\\myfiles*.txt
    :return: list of tuples consisting of:
        0 = Fully qualified path name of the file
        1 = Size of the file in bytes on the disk as an integer
        2 = Last date/time file was changed as a datetime.datetime value
        3 = Created date/time as a datetime.datetime value
        4 = Attribute list: A = Archived, R = ReadOnly, S = System, H = Hidden as a string
        Always returns a list, but that list may be empty.
    """
    xRet = list()
    xFiles = glob.iglob(cSkel)
    for cFile in xFiles:
        if _ver3x:
            xStat = os.stat(cFile, follow_symlinks=False)
        else:
            xStat = os.stat(cFile)
        if xStat is not None:
            nByteSize = xStat.st_size
            nModSeconds = xStat.st_mtime
            nCreateSeconds = xStat.st_ctime
            tMod = datetime.datetime.fromtimestamp(nModSeconds)
            tCreate = datetime.datetime.fromtimestamp(nCreateSeconds)
            cAttribute = ""
            if _ver3x:
                xAttributes = xStat.st_file_attributes
                if xAttributes & stat.FILE_ATTRIBUTE_ARCHIVE:
                    cAttribute += "A"
                if xAttributes & stat.FILE_ATTRIBUTE_HIDDEN:
                    cAttribute += "H"
                if xAttributes & stat.FILE_ATTRIBUTE_READONLY:
                    cAttribute += "R"
                if xAttributes & stat.FILE_ATTRIBUTE_SYSTEM:
                    cAttribute += "S"
            xRet.append((cFile, nByteSize, tCreate, tMod, cAttribute))
    return xRet

def FILETOSTR(lcFile):
    """
    To emulate the VFP function this MUST be a type 'rb', since VFP does NO transformations in a
    FILETOSTR() function.  Note that in Python 3.x, this will return a bytes object, not a string.  In that
    case you'll need to convert to unicode with an encoding, based on what you expect is in the file.
    """
    try:
        lcReturn = open(lcFile, 'rb').read()
    except:  # yes, any error
        lcReturn = ''
    # if _ver3x:
    #     lcReturn = lcReturn.decode("UTF-8", "ignore")
    return lcReturn

def STRTOFILE(lcString, lcFile, lbAppend=False):
    """
    Added optional 3rd parm to cause the string contents to be appended to an existing file rather than creating a new file.
    Added trap for bad directory name, which causes this to return a False.
    """
    bReturn = True
    lxFile = None

    try:
        if not lbAppend:
            if _ver3x:
                if isinstance(lcString, str):
                    lxFile = open(lcFile, "w", -1)
                else:
                    lxFile = open(lcFile, "wb", -1)
            else:
                lxFile = open(lcFile, 'wb', -1)
        else:
            if _ver3x:
                if isinstance(lcString, str):
                    lxFile = open(lcFile, "a", -1)
                else:
                    lxFile = open(lcFile, "ab", -1)
            else:
                lxFile = open(lcFile, 'ab', -1)
    except:
        bReturn = False
    if bReturn:
        lxFile.write(lcString)
        lxFile.flush()
        os.fsync(lxFile.fileno())  # Added flushing. 07/19/2016. JSH.
        lxFile.close()

    return bReturn
def get_product_version(path):
    try:
        # Query the fixed info part of the file version resource
        info = win32api.GetFileVersionInfo(path, "\\")
        ms = info['ProductVersionMS']
        ls = info['ProductVersionLS']
        version = "UNKNOWN"
        # Combine major, minor, subminor, and revision numbers
        # HIWORD and LOWORD extract the 16-bit parts of the 32-bit values
        if _ver3x:
            try:
                # version = f"{win32api.HIWORD(ms)}.{win32api.LOWORD(ms)}.{win32api.HIWORD(ls)}.{win32api.LOWORD(ls)}"
                version = str(win32api.HIWORD(ms)) + "." + str(win32api.LOWORD(ms)) + "." + str(win32api.HIWORD(ls)) + "." + str(win32api.LOWORD(ls))

            except:
                pass
        else:
            version = "Obsolete 2x version"
        return version
    except Exception as e:
        return "Error: " + str(e)



if __name__ == '__main__':
    xArgs = sys.argv
    if len(xArgs) > 1:
        cSelectedPython = xArgs[1]
        print("Attempting to choose Python instance %s" % cSelectedPython)
    else:
        cSelectedPython = ""
        print("Launching Python Chooser with Python List")
    ChoosePython(cSelectedPython=cSelectedPython)
    sys.exit(0)



# cd CBToolsInstallDir
# IF EXIST "c:\python27" copy /Y codebasetools.pth c:\python27\lib\site-packages\codebasetools.pth
# IF EXIST "C:\python36" copy /Y codebasetools.pth c:\python36\lib\site-packages\codebasetools.pth
# IF EXIST "C:\python37" copy /Y codebasetools.pth c:\python37\lib\site-packages\codebasetools.pth
# IF EXIST "C:\python38" copy /Y codebasetools.pth c:\python38\lib\site-packages\codebasetools.pth
# IF EXIST "C:\python39" copy /Y codebasetools.pth c:\python39\lib\site-packages\codebasetools.pth
# IF EXIST "C:\python310" copy /Y codebasetools.pth c:\python310\lib\site-packages\codebasetools.pth
# IF EXIST "C:\python311" copy /Y codebasetools.pth c:\python311\lib\site-packages\codebasetools.pth
# IF EXIST "C:\python311_64" copy /Y codebasetools.pth c:\python311_64\lib\site-packages\codebasetools.pth
# IF EXIST "C:\python312" copy /Y codebasetools.pth c:\python312\lib\site-packages\codebasetools.pth
# cd codebasetools
# IF EXIST "C:\python27" (IF NOT EXIST "c:\python27\lib\site-packages\codebasetools" MD c:\python27\lib\site-packages\codebasetools)
# IF EXIST "C:\python36" (IF NOT EXIST "c:\python38\lib\site-packages\codebasetools" MD c:\python36\lib\site-packages\codebasetools)
# IF EXIST "C:\python37" (IF NOT EXIST "c:\python38\lib\site-packages\codebasetools" MD c:\python37\lib\site-packages\codebasetools)
# IF EXIST "C:\python38" (IF NOT EXIST "c:\python38\lib\site-packages\codebasetools" MD c:\python38\lib\site-packages\codebasetools)
# IF EXIST "C:\python39" (IF NOT EXIST "c:\python39\lib\site-packages\codebasetools" MD c:\python39\lib\site-packages\codebasetools)
# IF EXIST "C:\python310" (IF NOT EXIST "c:\python310\lib\site-packages\codebasetools" MD c:\python310\lib\site-packages\codebasetools)
# IF EXIST "C:\python311" (IF NOT EXIST "c:\python311\lib\site-packages\codebasetools" MD c:\python311\lib\site-packages\codebasetools)
# IF EXIST "C:\python311_64" (IF NOT EXIST "c:\python311_64\lib\site-packages\codebasetools" MD c:\python311_64\lib\site-packages\codebasetools)
# IF EXIST "C:\python312" (IF NOT EXIST "c:\python312\lib\site-packages\codebasetools" MD c:\python312\lib\site-packages\codebasetools)
# IF EXIST "C:\python27\lib\site-packages\codebasetools" COPY /Y *.* C:\python27\lib\site-packages\codebasetools
# IF EXIST "C:\python36\lib\site-packages\codebasetools" COPY /Y *.* C:\python36\lib\site-packages\codebasetools
# IF EXIST "C:\python37\lib\site-packages\codebasetools" COPY /Y *.* C:\python37\lib\site-packages\codebasetools
# IF EXIST "C:\python38\lib\site-packages\codebasetools" COPY /Y *.* C:\python38\lib\site-packages\codebasetools
# IF EXIST "C:\python39\lib\site-packages\codebasetools" COPY /Y *.* C:\python39\lib\site-packages\codebasetools
# IF EXIST "C:\python310\lib\site-packages\codebasetools" COPY /Y *.* C:\python310\lib\site-packages\codebasetools
# IF EXIST "C:\python311\lib\site-packages\codebasetools" COPY /Y *.* C:\python311\lib\site-packages\codebasetools
# IF EXIST "C:\python311_64\lib\site-packages\codebasetools" COPY /Y *.* C:\python311_64\lib\site-packages\codebasetools
# IF EXIST "C:\python312\lib\site-packages\codebasetools" COPY /Y *.* C:\python312\lib\site-packages\codebasetools
# IF EXIST "C:\python27" copy /Y c:\python27\lib\site-packages\codebasetools\c4dll.dll c:\python27\c4dll.dll
# IF EXIST "C:\python36" copy /Y c:\python36\lib\site-packages\codebasetools\c4dll.dll c:\python36\c4dll.dll
# IF EXIST "C:\python37" copy /Y c:\python37\lib\site-packages\codebasetools\c4dll.dll c:\python37\c4dll.dll
# IF EXIST "C:\python38" copy /Y c:\python38\lib\site-packages\codebasetools\c4dll.dll c:\python38\c4dll.dll
# IF EXIST "C:\python39" copy /Y c:\python39\lib\site-packages\codebasetools\c4dll.dll c:\python39\c4dll.dll
# IF EXIST "C:\python310" copy /Y c:\python310\lib\site-packages\codebasetools\c4dll.dll c:\python310\c4dll.dll
# IF EXIST "C:\python311" copy /Y c:\python311\lib\site-packages\codebasetools\c4dll.dll c:\python311\c4dll.dll
# IF EXIST "C:\python311_64" copy /Y c:\python311_64\lib\site-packages\codebasetools\c4dll64.dll c:\python311\cdll64.dll
# IF EXIST "C:\python312" copy /Y c:\python312\lib\site-packages\codebasetools\c4dll.dll c:\python312\c4dll.dll
# IF EXIST "c:\python27" copy /Y c:\BGRepositoriesGIT\mpss-python-scripts-distro\mpsscommon\LibXLLicenseInfo.TXT c:\python27\lib\site-packages\codebasetools
# IF EXIST "C:\python36" copy /Y c:\BGRepositoriesGIT\mpss-python-scripts-distro\mpsscommon\LibXLLicenseInfo.TXT c:\python36\lib\site-packages\codebasetools
# IF EXIST "C:\python37" copy /Y c:\BGRepositoriesGIT\mpss-python-scripts-distro\mpsscommon\LibXLLicenseInfo.TXT c:\python37\lib\site-packages\codebasetools
# IF EXIST "C:\python38" copy /Y c:\BGRepositoriesGIT\mpss-python-scripts-distro\mpsscommon\LibXLLicenseInfo.TXT c:\python38\lib\site-packages\codebasetools
# IF EXIST "C:\python39" copy /Y c:\BGRepositoriesGIT\mpss-python-scripts-distro\mpsscommon\LibXLLicenseInfo.TXT c:\python39\lib\site-packages\codebasetools
# IF EXIST "C:\python310" copy /Y c:\BGRepositoriesGIT\mpss-python-scripts-distro\mpsscommon\LibXLLicenseInfo.TXT c:\python310\lib\site-packages\codebasetools
# IF EXIST "C:\python311" copy /Y c:\BGRepositoriesGIT\mpss-python-scripts-distro\mpsscommon\LibXLLicenseInfo.TXT c:\python311\lib\site-packages\codebasetools
# IF EXIST "C:\python311_64" copy /Y c:\BGRepositoriesGIT\mpss-python-scripts-distro\mpsscommon\LibXLLicenseInfo.TXT c:\python311_64\lib\site-packages\codebasetools
# IF EXIST "C:\python312" copy /Y c:\BGRepositoriesGIT\mpss-python-scripts-distro\mpsscommon\LibXLLicenseInfo.TXT c:\python312\lib\site-packages\codebasetools
# cd ..
# cd ..
# ECHO DONE