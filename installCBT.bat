cd CBToolsInstallDir
IF EXIST "c:\python27" copy /Y codebasetools.pth c:\python27\lib\site-packages\codebasetools.pth
IF EXIST "C:\python36" copy /Y codebasetools.pth c:\python36\lib\site-packages\codebasetools.pth
IF EXIST "C:\python37" copy /Y codebasetools.pth c:\python37\lib\site-packages\codebasetools.pth
IF EXIST "C:\python38" copy /Y codebasetools.pth c:\python38\lib\site-packages\codebasetools.pth
IF EXIST "C:\python39" copy /Y codebasetools.pth c:\python39\lib\site-packages\codebasetools.pth
IF EXIST "C:\python310" copy /Y codebasetools.pth c:\python310\lib\site-packages\codebasetools.pth
IF EXIST "C:\python311" copy /Y codebasetools.pth c:\python311\lib\site-packages\codebasetools.pth
IF EXIST "C:\python311_64" copy /Y codebasetools.pth c:\python311_64\lib\site-packages\codebasetools.pth
IF EXIST "C:\python312" copy /Y codebasetools.pth c:\python312\lib\site-packages\codebasetools.pth
cd codebasetools
IF EXIST "C:\python27" (IF NOT EXIST "c:\python27\lib\site-packages\codebasetools" MD c:\python27\lib\site-packages\codebasetools)
IF EXIST "C:\python36" (IF NOT EXIST "c:\python38\lib\site-packages\codebasetools" MD c:\python36\lib\site-packages\codebasetools)
IF EXIST "C:\python37" (IF NOT EXIST "c:\python38\lib\site-packages\codebasetools" MD c:\python37\lib\site-packages\codebasetools)
IF EXIST "C:\python38" (IF NOT EXIST "c:\python38\lib\site-packages\codebasetools" MD c:\python38\lib\site-packages\codebasetools)
IF EXIST "C:\python39" (IF NOT EXIST "c:\python39\lib\site-packages\codebasetools" MD c:\python39\lib\site-packages\codebasetools)
IF EXIST "C:\python310" (IF NOT EXIST "c:\python310\lib\site-packages\codebasetools" MD c:\python310\lib\site-packages\codebasetools)
IF EXIST "C:\python311" (IF NOT EXIST "c:\python311\lib\site-packages\codebasetools" MD c:\python311\lib\site-packages\codebasetools)
IF EXIST "C:\python311_64" (IF NOT EXIST "c:\python311_64\lib\site-packages\codebasetools" MD c:\python311_64\lib\site-packages\codebasetools)
IF EXIST "C:\python312" (IF NOT EXIST "c:\python312\lib\site-packages\codebasetools" MD c:\python312\lib\site-packages\codebasetools)
IF EXIST "C:\python27\lib\site-packages\codebasetools" COPY /Y *.* C:\python27\lib\site-packages\codebasetools
IF EXIST "C:\python36\lib\site-packages\codebasetools" COPY /Y *.* C:\python36\lib\site-packages\codebasetools
IF EXIST "C:\python37\lib\site-packages\codebasetools" COPY /Y *.* C:\python37\lib\site-packages\codebasetools
IF EXIST "C:\python38\lib\site-packages\codebasetools" COPY /Y *.* C:\python38\lib\site-packages\codebasetools
IF EXIST "C:\python39\lib\site-packages\codebasetools" COPY /Y *.* C:\python39\lib\site-packages\codebasetools
IF EXIST "C:\python310\lib\site-packages\codebasetools" COPY /Y *.* C:\python310\lib\site-packages\codebasetools
IF EXIST "C:\python311\lib\site-packages\codebasetools" COPY /Y *.* C:\python311\lib\site-packages\codebasetools
IF EXIST "C:\python311_64\lib\site-packages\codebasetools" COPY /Y *.* C:\python311_64\lib\site-packages\codebasetools
IF EXIST "C:\python312\lib\site-packages\codebasetools" COPY /Y *.* C:\python312\lib\site-packages\codebasetools
IF EXIST "C:\python27" copy /Y c:\python27\lib\site-packages\codebasetools\c4dll.dll c:\python27\c4dll.dll
IF EXIST "C:\python36" copy /Y c:\python36\lib\site-packages\codebasetools\c4dll.dll c:\python36\c4dll.dll
IF EXIST "C:\python37" copy /Y c:\python37\lib\site-packages\codebasetools\c4dll.dll c:\python37\c4dll.dll
IF EXIST "C:\python38" copy /Y c:\python38\lib\site-packages\codebasetools\c4dll.dll c:\python38\c4dll.dll
IF EXIST "C:\python39" copy /Y c:\python39\lib\site-packages\codebasetools\c4dll.dll c:\python39\c4dll.dll
IF EXIST "C:\python310" copy /Y c:\python310\lib\site-packages\codebasetools\c4dll.dll c:\python310\c4dll.dll
IF EXIST "C:\python311" copy /Y c:\python311\lib\site-packages\codebasetools\c4dll.dll c:\python311\c4dll.dll
IF EXIST "C:\python311_64" copy /Y c:\python311_64\lib\site-packages\codebasetools\c4dll64.dll c:\python311\cdll64.dll
IF EXIST "C:\python312" copy /Y c:\python312\lib\site-packages\codebasetools\c4dll.dll c:\python312\c4dll.dll
IF EXIST "c:\python27" copy /Y c:\BGRepositoriesGIT\mpss-python-scripts-distro\mpsscommon\LibXLLicenseInfo.TXT c:\python27\lib\site-packages\codebasetools
IF EXIST "C:\python36" copy /Y c:\BGRepositoriesGIT\mpss-python-scripts-distro\mpsscommon\LibXLLicenseInfo.TXT c:\python36\lib\site-packages\codebasetools
IF EXIST "C:\python37" copy /Y c:\BGRepositoriesGIT\mpss-python-scripts-distro\mpsscommon\LibXLLicenseInfo.TXT c:\python37\lib\site-packages\codebasetools
IF EXIST "C:\python38" copy /Y c:\BGRepositoriesGIT\mpss-python-scripts-distro\mpsscommon\LibXLLicenseInfo.TXT c:\python38\lib\site-packages\codebasetools
IF EXIST "C:\python39" copy /Y c:\BGRepositoriesGIT\mpss-python-scripts-distro\mpsscommon\LibXLLicenseInfo.TXT c:\python39\lib\site-packages\codebasetools
IF EXIST "C:\python310" copy /Y c:\BGRepositoriesGIT\mpss-python-scripts-distro\mpsscommon\LibXLLicenseInfo.TXT c:\python310\lib\site-packages\codebasetools
IF EXIST "C:\python311" copy /Y c:\BGRepositoriesGIT\mpss-python-scripts-distro\mpsscommon\LibXLLicenseInfo.TXT c:\python311\lib\site-packages\codebasetools
IF EXIST "C:\python311_64" copy /Y c:\BGRepositoriesGIT\mpss-python-scripts-distro\mpsscommon\LibXLLicenseInfo.TXT c:\python311_64\lib\site-packages\codebasetools
IF EXIST "C:\python312" copy /Y c:\BGRepositoriesGIT\mpss-python-scripts-distro\mpsscommon\LibXLLicenseInfo.TXT c:\python312\lib\site-packages\codebasetools
cd ..
cd ..
ECHO DONE
