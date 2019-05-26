# coding=utf-8
# 
from xlrd import open_workbook
from functools import partial
from itertools import takewhile, chain
from cmbhk.utility import getCurrentDirectory, getStartRow, getCustodian
from utils.excel import worksheetToLines, rowToList
import logging
logger = logging.getLogger(__name__)



def readHolding(ws, startRow):
	"""
	[Worksheet] ws, [Int] startRow => [Iterable] rows
	
	Read the Excel worksheet containing the holdings, return an iterable object
	on the list of holding positions. Each position is a dictionary.
	"""
	headers = readHeaders(ws, startRow)
	position = lambda headers, values: dict(zip(headers, values))
	return map(partial(position, headers)
			  , takewhile(isHolding
						 , worksheetToLines(ws
								  		   , getStartRow()+1
								  		   , len(headers))))



def readHeaders(ws, startRow):
	"""
	[Worksheet] ws, [Int] startRow => [List] headers
	"""
	toString = lambda s: str(s)
	firstLine = lambda s: s.split('\n')[0].strip()
	nonEmptyString = lambda s: s != ''

	return list(takewhile(nonEmptyString
						 , map(firstLine
						  	  , map(toString
						  	  	   , rowToList(ws, startRow)))))



def stripIfString(x):
	if isinstance(x, str):
		return x.strip()
	else:
		return x



def isHolding(lineItems):
	"""
	[List] lineItems => [Boolean] is this a holding line

	A line from worksheetToLines() is a list of items, the function tells 
	whether that list represents a holding positions.
	"""
	try:
		return lineItems[0] != ''

	except IndexError:
		return False



if __name__ == '__main__':
	import logging.config
	logging.config.fileConfig('logging.config', disable_existing_loggers=False)

	from os.path import join
	inputFile = join(getCurrentDirectory(), 'samples', 'holding _ 16032017.xlsx')
	wb = open_workbook(inputFile)
	ws = wb.sheet_by_index(0)
	print(readHeaders(ws, getStartRow()))
	holding = readHolding(ws, getStartRow())
	for h in holding:
		print(h)