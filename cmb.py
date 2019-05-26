# coding=utf-8
# 
from xlrd import open_workbook
from functools import partial
from itertools import takewhile, dropwhile, chain
from cmbhk.utility import getCurrentDirectory, getStartRow, getCustodian
from utils.excel import worksheetToLines, rowToList
from utils.iter import head, firstOf
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
			  , takewhile(firstCellNotEmpty
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



def readCash(ws, startRow):
	"""
	[Worksheet] ws, [Int] startRow => [Tuple] (currency, amount)

	Return tuple like ('HKD', 12345.67)
	"""
	hasClosingBalance = lambda L: True if any(isClosingBalance(x) for x in L) else False
	lineItems = firstOf(hasClosingBalance, worksheetToLines(ws, getStartRow()+1))
	if lineItems == None:
		raise ValueError('readCash(): cash line is None')

	isFloat = lambda x: True if isinstance(x, float) else False
	amount = firstOf(isFloat, lineItems)
	currencyString = firstOf(isClosingBalance, lineItems)
	if amount == None or currencyString == None:
		raise ValueError('{0}, {1}'.format(currencyString, amount))

	return (currencyString.strip()[-4:-1], amount)

	

def isClosingBalance(s):
	if isinstance(s, str) and s.startswith('Closing Balance'):
		return True
	else:
		return False



def firstCellNotEmpty(lineItems):
	"""
	[List] lineItems => [Boolean] is the first item an empty string?

	If the list has a first cell and it is an empty string, return False;
	else return True.
	"""
	try:
		return lineItems[0] != ''

	except IndexError:
		return False



def genevaPosition(portId, date, position):
	"""
	[String] portId, [String] date, [Dictionary] position => 
		[Dictionary] gPosition

	A Geneva position is a dictionary object that has the following list
	of keys:

	portfolio|custodian|date|geneva_investment_id|ISIN|bloomberg_figi|name
	|currency|quantity
	
	"""
	genevaPos = {}
	genevaPos['portfolio'] = portId
	genevaPos['custodian'] = getCustodian()
	genevaPos['date'] = date
	genevaPos['name'] = position['Securities Name']
	genevaPos['currency'] = position['Ccy']
	genevaPos['quantity'] = position['Traded Quantity']
	genevaPos['geneva_investment_id'] = ''
	genevaPos['ISIN'] = position['Securities Identifier']
	genevaPos['bloomberg_figi'] = ''
	
	return genevaPos



if __name__ == '__main__':
	import logging.config
	logging.config.fileConfig('logging.config', disable_existing_loggers=False)

	from os.path import join
	inputFile = join(getCurrentDirectory(), 'samples', 'holding _ 16032017.xlsx')
	wb = open_workbook(inputFile)
	ws = wb.sheet_by_index(0)

	gPositions = map(partial(genevaPosition, '40017', '2017-03-16') 
					, readHolding(ws, getStartRow()))
	for x in gPositions:
		print(x)


	inputFile = join(getCurrentDirectory(), 'samples', 'cash _ 16032017.xlsx')
	wb = open_workbook(inputFile)
	ws = wb.sheet_by_index(0)
	print(readCash(ws, getStartRow()))