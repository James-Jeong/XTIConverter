# -*- coding: utf-8 -*-
# XTIConverter.py
# author: jamesj

from header.XTIConverter_h import *

# ----------------------------------------------------------------------------------------- #
# MAIN FUNCTION

# @fn main
# @brief XSLX To File Converter main function
if __name__ == '__main__':
# ----------------------------------------------------------------------------------------- #
# 1) Get Parameters
	xlsxPath = None
	sheetName = None
	iniPath = None

# 1-1) Check program mode
	xlsxPath, sheetName, iniPath = checkMode(sys.argv)

	print("\n@ XLSX Path: [ " + xlsxPath + " ]")
	print("@ Sheet Name: [ " + sheetName + " ]")
	print("@ INI Path: [ " + iniPath + " ]")

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

