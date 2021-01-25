# -*- coding: utf-8 -*-
# Logger_h.py
# author: jamesj

import logging

class Logger:
	def __init__(self, name, level, msgFmt, dateFmt):
		self.log = logging.getLogger(name)
		self.log.propagate = True
		self.formatter = logging.Formatter(msgFmt, dateFmt)
		self.levels = {
			"DEBUG" : logging.DEBUG,
			"INFO" : logging.INFO,
			"WARNING" : logging.WARNING,
			"ERROR" : logging.ERROR,
			"CRITICAL" : logging.CRITICAL
		}
		self.log.setLevel(level=self.levels[level])

	def getStreamHandler(self):
		streamHandler = logging.StreamHandler()
		#streamHandler.setLevel(level=self.levels[level])
		streamHandler.setFormatter(self.formatter)
		self.log.addHandler(streamHandler)
		return self.log

