# -*- coding: utf-8 -*-
# XTIConverter.py
# author: jamesj

from header.XTIConverter_h import *

# ----------------------------------------------------------------------------------------- #
# MAIN FUNCTION

# @fn main
# @brief XSLX To INI File Converter main function
if __name__ == '__main__':
# ----------------------------------------------------------------------------------------- #
# 0) Set logging
	log.setLogging(Logger.INFO_LEVEL, "%(asctime)s:%(module)s:%(levelname)s:%(message)s", "%Y-%m-%d %H:%M:%S")

# ----------------------------------------------------------------------------------------- #
# 1) Get Parameters
	xlsxPath = None
	sheetName = None
	iniPath = None

# 1-1) Check program mode
	xlsxPath, sheetName, iniPath = checkMode(sys.argv)

	log.logInfo("@ XLSX Path: [ " + xlsxPath + " ]")
	log.logInfo("@ Sheet Name: [ " + sheetName + " ]")
	log.logInfo("@ INI Path: [ " + iniPath + " ]")

# ----------------------------------------------------------------------------------------- #
# 2) Check xlsx path & ini path
	checkXlsx(xlsxPath)
	checkIni(iniPath)

# ----------------------------------------------------------------------------------------- #
# 3) Load xlsx
	totalData = loadXlsx(xlsxPath, sheetName)

# ----------------------------------------------------------------------------------------- #
# 4) Write ini file
	writeIni(iniPath, totalData)

