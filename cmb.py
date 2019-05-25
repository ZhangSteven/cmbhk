# coding=utf-8
# 
from xlrd import open_workbook
from functools import partial
from itertools import takewhile, chain
from cmbhk.utility import getCurrentDirectory, getStartRow, getCustodian
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
	firstLine = lambda s: s.split('\n')[0].strip()
	return list(map(firstLine
			       , takewhile(lambda x: x != ''
						  	  , map(stripIfString
						  	   	   , map(partial(cellValue, ws, startRow)
						  	   	    	, range(ws.ncols))))))



def stripIfString(x):
	if isinstance(x, str):
		return x.strip()
	else:
		return x



def cellValue(ws, row, column):
	return ws.cell_value(row, column)



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



def worksheetToLines(ws, startRow=0, numItems=None):
	"""
	[Worksheet] ws, [Int] startRow, [Int] numItems => [Iterable] a list of lines
	Where,

	ws: worksheet
	startRow: the starting line to read from
	numItems: number of columns to read each line

	Each time the generator yields a line, which is a list of values from column 0
	up to the number of columns to read
	"""
	row = startRow
	if numItems == None:
		numItems = ws.ncols

	while (row < ws.nrows):
		yield list(map(partial(cellValue, ws, row), range(numItems)))
		row = row + 1




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