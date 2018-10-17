# Copyright (C) 2010-2018 by M-P Systems Services, Inc.,
# PMB 136, 1631 NE Broadway, Portland, OR 97232.
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU Lesser General Public License as published by
# the Free Software Foundation, version 3 of the License.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY
# or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU Lesser General
# Public License for more details.
#
# You should have received a copy of the GNU Lesser General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

from __future__ import print_function, absolute_import
import logging
import logging.config
from win32com.server.exception import COMException
import os
import sys
import time
import inspect
import traceback
import MPSSBaseTools as mTools

__author__ = "J. Heuer, J. McRae"


class COMprops (object):
    """ Basic COM object class used as parent or co-parent for all Python-based COM object to be used by MPSS VFP Web Proc type
    applications.  The properties cErrorFilePath, cDataDirectory, cClientDataDirectory, cTempFilePath, and others
    that may be added from time to time should be set by the caller.  Normally the VFP Class MPPythonCOMWrapper should
    be instantiated and then CREATEPYTHONOBJECT() used to create the Python COM object.  It will then take on the job of
    setting the necessary properties. """
    _public_methods_ = [ 'set_property', 'configure',
                         'get_property', 'dump_contents', 'start_logging', 'set_error_reporting', 'known_error' ]
    _public_attrs_ = ['cTempFilePath', 'cErrorMessage', 'nErrorNumber', 'name']
    # Use pythoncom.CreateGuid()" to make a new clsid
    # _reg_clsid_ = '{A6706465-C80D-4F40-9D04-E36EAC9F94B5}'
    # The above properties should be set by the implementation subclass of this class

    def __init__(self):
        self.logger = None
        self.cErrorMessage = ""
        self.cErrorFilePath = ""
        self.cClientDataDirectory = ""
        self.cTempFilePath = ""
        self.nErrorNumber = 0
        self.name = "Base COM Object"
        self.logpath = ""
        self.debug = False
        
    def set_property(self, name, value):
        setattr(self, name, value)

    def get_property(self, name):
        return getattr(self, name, "NOT FOUND")

    def known_error(self, thing):  # Use as template for any methods to be called from COM.
        """ The reason the entire method body is wrapped in a try-catch is that exceptionhook does NOT
        work in COM server situations.  The entire trace-back error is returned as the OLE error, which
        works OK in VFP, but it is impossible to trigger a complete error trace dump file without the
        ability to trap the exception with an exceptionhook method.  By wrapping everything and raising
        the COMException, we detect any unhandled exceptions, write out a detailed debug report, and then
        raise an exception which can be detected and reported on by the VFP caller client."""
        myLocal1 = 2359
        try:
            # Main body of the method to be called as a COM object method
            myLocal1 = 234234
            myLocal2 = 23412398
            myTest = self.known_error2()
        except:  # bare exception that stores the error condition and raises the necessary COM error
            lcType, lcValue, lxTrace = sys.exc_info()
            self.local_except_handler(lcType, lcValue, lxTrace)
            raise COMException(description = ("Python Error Occurred: " + repr(lcType)))
        return myLocal1

    def known_error2(self):  # notused
        myLocal3 = 2344
        myLocal4 = 4654
        myBadDiv = 34924/0.0
        return True

    def configure(self, array):
        for (kw, value) in array:
            self.set_property(kw, value)

    def dump_contents(self):
        return self.__dict__.items()

    def output_error_info(self):  # notused
        lcType, lxValue, lxTrace = sys.exc_info()
        return self.local_except_handler(lcType, lxValue, lxTrace)

    def local_except_handler(self, cType, xValue, xTraceback):
        lcBaseErrFile = "PYCOMERR_" + mTools.strRand(8, True) + ".ERR"
        lcErrFile = mTools.ADDBS(self.cErrorFilePath) + lcBaseErrFile
        print(lcErrFile)
        lbFileOK = False

        loHndl = None
        try:
            loHndl = open(lcErrFile, "w", -1)
            lbFileOK = True
        except IOError:
            lbFileOK = False
        if not lbFileOK:
            loHndl = open(lcBaseErrFile, "w", -1)
            # If the above fails, it's BAD
            
        loHndl.write("COM Error Report: ")
        chgTimex = time.localtime()
        lcChgTime = time.strftime("%Y-%m-%d %H:%M:%S", chgTimex)
        loHndl.write(lcChgTime + "\n\n")
        loHndl.write(repr(cType) + "\n")
        loHndl.write(repr(xValue) + "\n\nEXCEPTION TRACE:\n")

        traceback.print_exception(cType, xValue, xTraceback, file=loHndl)

        loHndl.write("\n\nCONTENTS DUMP:\n")
        xDump = self.dump_contents()
        for jj in xDump:
            loHndl.write(repr(jj) + "\n")

        loHndl.write("\n\nVARIABLES DUMP:\n")
            
        workTraceback = xTraceback
        if workTraceback is None:
            return False
        
        xFrame = workTraceback.tb_frame
        loHndl.write("GLOBALS:\n")
        for jj in xFrame.f_globals:
            if repr(jj) != "'__builtins__'":
                loHndl.write(repr(jj) + " = " + repr(xFrame.f_globals[jj]) + "\n")

        loHndl.write("\n")        
        while workTraceback is not None:
            xFrame = workTraceback.tb_frame            
            lcName = xFrame.f_code.co_filename
            loHndl.write("\nMODULE NAME  : " + lcName + "\n")
            lcObjName = xFrame.f_code.co_name
            loHndl.write(  "FUNCTION NAME: " + lcObjName + "\n")
            loHndl.write("LOCALS:\n")
            for kk in xFrame.f_locals:
                loHndl.write(repr(kk) + " = " + repr(xFrame.f_locals[kk]) + "\n")
            workTraceback = workTraceback.tb_next
        
        loHndl.close()
        loHndl = None
        return True

    def set_error_reporting(self, lcErrFilePath):
        """ Sets exception handling to write to the specified error log directory for debugging purposes. """
        self.cErrorFilePath = lcErrFilePath
        # sys.excepthook = self.local_except_handler
        return True

    def start_logging(self):
        """ Deprecated.  Use set_error_reporting() instead. """
        self.logger = None
        if mTools.confirmPath(self.logpath):  # expecting MPSS standard path
            LOG_FILENAME = os.path.normpath(
                os.path.join(self.logpath, mTools.strRand(8, False) + ".ERR"))
        else:
            # fallback to root directory
            LOG_FILENAME = os.path.normpath(os.path.join("C:\\", "PYCOM_DEFAULT.ERR"))

        # print LOG_FILENAME
        # delay=True keeps prevents file i/o until first emit
        # there won't be any if no logging occurs -- don't want disk files then
        h = logging.FileHandler(filename=LOG_FILENAME, delay=True)
        h.level = logging.DEBUG
        formatter = logging.Formatter('%(name)-12s: %(levelname)-8s %(message)s')
        # tell the handler to use this format
        h.setFormatter(formatter)
        self.logger = logging.Logger("MPSS_pycom")
        self.logger.addHandler(h)
        self.debug = True


if __name__ == "__main__":
    # print "Registering COM server..."
    # import win32com.server.register
    # win32com.server.register.UseCommandLine(COMprops)
    print(mTools.strRand(8, True))
    xo = COMprops()
    xo.set_error_reporting("e:\\temp")
    # xtest = xo.known_error(3)

