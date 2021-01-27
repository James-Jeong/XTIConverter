# -*- coding: utf-8 -*-
# XTIConverter_h.py
# author: jamesj

import os
import sys
from openpyxl import load_workbook
import configparser

from .Logger_h import Logger

# ----------------------------------------------------------------------------------------- #
# PREDEFINES
log = None

# ----------------------------------------------------------------------------------------- #
# FUNCTIONS

# @fn existProgram()
# @brief 프로그램을 종료하는 함수
def exitProgram():
	sys.exit()

# @fn checkMode(argv)
# @brief 프로그램 모드를 확인하는 함수
def checkMode(argv):
	global log
	xlsxPath = None
	sheetName = None
	iniPath = None
	argvLen = len(argv)

	if argvLen == 2:
		converterConfigPath = argv[1]
		if not os.path.exists(converterConfigPath) or not os.path.isfile(converterConfigPath):
			exitProgram()

		converterConfig = configparser.ConfigParser()
		converterConfig.optionxform = lambda option: option # Preserve case for letters
		converterConfig.read(converterConfigPath)
	
		# PATH section 이 없으면 종료	
		if not converterConfig.has_section("PATH"):
			exitProgram()
		# XLSX section 이 없으면 종료	
		if not converterConfig.has_section("XLSX"):
			exitProgram()
		# LOG section 이 없으면 종료	
		if not converterConfig.has_section("LOG"):
			exitProgram()

		# LOG section 에서 LEVEL key 의 value 를 조회
		logLevel = converterConfig.get("LOG", "LEVEL")
		# Log level 이 정의되지 않으면 종료
		if logLevel is None or len(logLevel) == 0:
			exitProgram()

		# Log 형식 지정
		log = Logger("logger", logLevel, "%(asctime)s [%(levelname)s] %(message)s", "[ %Y-%m-%d %H:%M:%S ]")
		log = log.getStreamHandler()

		# PATH section 에서 XLSX_PATH key 의 value 를 조회
		xlsxPath = converterConfig.get("PATH", "XLSX_PATH")
		if xlsxPath is None or len(xlsxPath) == 0:
			log.info("! Fail to load the option \"XLSX_PATH\" in the section \"XLSX\". (configPath" + converterConfigPath + ")\n")
			exitProgram()

		# PATH section 에서 INI_PATH key 의 value 를 조회
		iniPath = converterConfig.get("PATH", "INI_PATH")
		if iniPath is None or len(iniPath) == 0:
			log.warn("! Fail to load the option \"INI_PATH\" in the section \"PATH\". (configpath={}".format(converterConfigPath))
			exitProgram()

		# XLSX section 에서 SHEET_NAME key 의 value 를 조회
		sheetName = converterConfig.get("XLSX", "SHEET_NAME")
		if sheetName is None or len(sheetName) == 0:
			log.warn("! Fail to load the option \"SHEET_NAME\" in the section \"PATH\". (configPath={}".format(converterConfigPath))
			exitProgram()

		log.info("- XLSX Path: [ {} ]".format(xlsxPath))
		log.info("- Sheet Name: [ {} ]".format(sheetName))
		log.info("- INI Path: [ {} ]".format(iniPath))
		log.info("- Log level: [ {} ]".format(logLevel))
		log.info("Loading config...(OK)")
	else:
		log.info("Parameter error.")
		log.info("	argv[0]: XTIConverter.py")
		log.info("	argv[1]: {config path}")

	return xlsxPath, sheetName, iniPath

# @fn checkXlsx(xlsxPath)
# @brief 지정한 XLSX 경로 존재 여부와 확장자를 검사하는 함수
def checkXlsx(xlsxPath):
	if not os.path.exists(xlsxPath) or not os.path.isfile(xlsxPath):
		log.warn("Unknown XLSX Path.")
		exitProgram()

	xlsxName, xlsxExtension = os.path.splitext(xlsxPath)
	if xlsxExtension != ".xlsx":
	        log.warn("XLSX File type is wrong. (ext={}".format(xlsxExtension))
        	exitProgram()

# @fn checkIni(iniPath)
# @brief 지정한 INI 파일의 확장자를 검사하는 함수
def checkIni(iniPath):
	iniName, iniExtension = os.path.splitext(iniPath)
	if iniExtension != ".ini":
	        log.warn("INI File type is wrong. (ext={}".format(iniExtension))
	        exitProgram()

# @fn loadXlsx(xlsxPath, sheetName)
# @brief 지정한 XLSX 파일 로드하여 저장된 값들을 반환하는 함수
def loadXlsx(xlsxPath, sheetName):
	load_wb = load_workbook(xlsxPath, data_only = True)
	load_ws = load_wb[sheetName]

	totalData = []
	for row in load_ws.rows:
		row_value = []
		for cell in row:
			if cell.value is None:
				break
			row_value.append(str(cell.value).strip())
		totalData.append(row_value)
	
	if len(totalData) <= 1:
		log.info("Loading xlsx...(FAIL): The xlsx file is empty.")
		exitProgram()
	else:
		for row in totalData:
			if len(row) == 0:
				log.debug()
				continue
			if len(row) == 1:
				log.debug("[ {} ]".format(row[0]))
			elif len(row) == 2:
				log.debug("	- {}: {}".format(row[0], row[1]))
			else:
				log.debug(row)

	log.info("Loading xlsx...(OK)")
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
		log.warn("Writing ini...(FAIL): Not found any section.")
		exitProgram()

	with open(iniPath, 'w', encoding='utf8') as configfile:
		config.write(configfile)

	if os.path.exists(iniPath) and os.path.isfile(iniPath):
		log.info("Writing ini...(OK)")
	else:
		log.warn("Writing ini...(FAIL): Not found the ini file.")

