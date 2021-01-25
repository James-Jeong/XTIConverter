# -*- coding: utf-8 -*-
# Logger_h.py
# author: jamesj

import logging

class Logger:
	INFO_LEVEL = logging.INFO
	DEBUG_LEVEL = logging.DEBUG
	WARNING_LEVEL = logging.warning

	logger = logging.getLogger(__name__)

	def logInfo(self, msg):
		self.logger.info(msg)

	def logDebug(self, msg):
		self.logger.debug(msg)

	def logWarning(self, msg):
		self.logger.warning(msg)

	def setLogging(self, level, msgFmt, dateFmt):
		formatter = logging.Formatter(msgFmt, dateFmt)
		streamHandler = logging.StreamHandler()
		streamHandler.setLevel(level)
		streamHandler.setFormatter(formatter)
		self.logger.addHandler(streamHandler)

