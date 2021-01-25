# -*- coding: utf-8 -*-
# XTIConverter_h.py
# author: jamesj

import os
import sys
from openpyxl import load_workbook
import configparser

# ----------------------------------------------------------------------------------------- #
# FUNCTIONS

# @fn existProgram()
# @brief 프로그램을 종료하는 함수
def exitProgram():
	sys.exit()

# @fn checkMode(argv)
# @brief 프로그램 모드를 확인하는 함수
# 1. Config mode
# 2. Paramter mode
def checkMode(argv):
	xlsxPath = None
	sheetName = None
	iniPath = None
	argvLen = len(argv)

	# 1) Config mode
	if argvLen == 2:
		print("\n@ Config mode")
		print("- Loading config...")

		converterConfigPath = argv[1]
		if not os.path.exists(converterConfigPath) or not os.path.isfile(converterConfigPath):
			print("! Fail to load the program's config file. (name=" + converterConfigPath + ")\n")
			exitProgram()

		converterConfig = configparser.ConfigParser()
		converterConfig.optionxform = lambda option: option # Preserve case for letters
		converterConfig.read(converterConfigPath)
	
		if not converterConfig.has_section("PATH"):
			print("! Fail to load the section \"PATH\". (name=" + converterConfigPath + ")\n")
			exitProgram()
		if not converterConfig.has_section("XLSX"):
			print("! Fail to load the section \"XLSX\". (name=" + converterConfigPath + ")\n")
			exitProgram()

		xlsxPath = converterConfig.get("PATH", "XLSX_PATH")
		if xlsxPath is None or len(xlsxPath) == 0:
			print("! Fail to load the option \"XLSX_PATH\" in the section \"PATH\". (name=" + converterConfigPath + ")\n")
			exitProgram()

		iniPath = converterConfig.get("PATH", "INI_PATH")
		if iniPath is None or len(iniPath) == 0:
			print("! Fail to load the option \"INI_PATH\" in the section \"PATH\". (name=" + converterConfigPath + ")\n")
			exitProgram()

		sheetName = converterConfig.get("XLSX", "SHEET_NAME")
		if sheetName is None or len(sheetName) == 0:
			print("! Fail to load the option \"SHEET_NAME\" in the section \"XLSX\". (name=" + converterConfigPath + ")\n")
			exitProgram()

	# 2) Parameter mode
	elif argvLen == 4:
		print("\n@ Parameter mode")
		print("- Loading parameters...")

		xlsxPath = argv[1]
		if len(xlsxPath) == 0:
			print("! Fail to load the xslx path.\n")
			exitProgram()

		sheetName = argv[2]
		if len(sheetName) == 0:
			print("! Fail to load the sheet name.\n")
			exitProgram()

		iniPath = argv[3]
		if len(iniPath) == 0:
			print("! Fail to load the ini path.\n")
			exitProgram()

	else:
		print("\n! Parameter error.")
		print("argv[0]: XTIConverter.py\n")
		print("1) Config mode")
		print("argv[1]: {config path}\n")
		print("2) Parameter mode")
		print("argv[1]: {xlsx path}")
		print("argv[2]: {xlsx sheet name}")
		print("argv[3]: {ini path}\n")
		exitProgram()

	return xlsxPath, sheetName, iniPath

# @fn checkXlsx(xlsxPath)
# @brief 지정한 XLSX 경로 존재 여부와 확장자를 검사하는 함수
def checkXlsx(xlsxPath):
	if not os.path.exists(xlsxPath) or not os.path.isfile(xlsxPath):
		print("! Unknown XLSX Path.\n")
		exitProgram()

	xlsxName, xlsxExtension = os.path.splitext(xlsxPath)
	if xlsxExtension != ".xlsx":
	        print("! XLSX File type is wrong. (ext=" + xlsxExtension + ")\n")
        	exitProgram()

# @fn checkIni(iniPath)
# @brief 지정한 INI 파일의 확장자를 검사하는 함수
def checkIni(iniPath):
	iniName, iniExtension = os.path.splitext(iniPath)
	if iniExtension != ".ini":
	        print("! INI File type is wrong. (ext=" + iniExtension + ")\n")
	        exitProgram()

# @fn loadXlsx(xlsxPath, sheetName)
# @brief 지정한 XLSX 파일 로드하여 저장된 값들을 반환하는 함수
def loadXlsx(xlsxPath, sheetName):
	load_wb = load_workbook(xlsxPath, data_only = True)
	load_ws = load_wb[sheetName]

	# 3-1) Parsing
	print("\n- Loading xlsx...\n")
	totalData = []
	for row in load_ws.rows:
		row_value = []
		for cell in row:
			if cell.value is None:
				break
			row_value.append(str(cell.value).strip())
		totalData.append(row_value)

	# 3-2) Check result
	if len(totalData) <= 1:
		print("! Fail to load. The file is empty.\n")
		exitProgram()
	else:
		for row in totalData:
			if len(row) == 0:
				print()
				continue
			if len(row) == 1:
				print("- [" + row[0] + "]")
			elif len(row) == 2:
				print("	- " + row[0] + ": " + row[1])
			else:
				print(row)
		print("\n- Success to load.")

	return totalData

# @fn writeIni(iniPath, totalData)
# @brief 지정한 INI 경로와 저장된 값들을 통해 새로운 INI 파일을 생성하는 함수
def writeIni(iniPath, totalData):
	config = configparser.ConfigParser()
	config.optionxform = lambda option: option # Preserve case for letters

	curSection = None
	for row in totalData:
		row_value = []
		for cell in row:
			row_value.append(cell)
		# 1. Row 에 section 이 있는 경우
		if len(row_value) == 1:
			curSection = row_value[0]
			config.add_section(curSection)
		# 2. Row 에 key, value 가 있는 경우
		elif len(row_value) == 2 and curSection is not None:
			config.set(curSection, row_value[0], row_value[1])
		# 3. Row 가 비었거나 ini 형식과 일치하지 않는 경우
		else:
			curSection = None

	if curSection is None:
		print("! Fail. Not found any section.\n")
		exitProgram()

	with open(iniPath, 'w', encoding='utf8') as configfile:
		config.write(configfile)

	if os.path.exists(iniPath):
		print("\n@ Done.\n")
	else:
		print("\n! Fail.\n")

