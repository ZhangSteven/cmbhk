# coding=utf-8
# 
from xlrd import open_workbook
from functools import partial
from itertools import takewhile, chain
from operator import getitem
from os.path import join
from hsbc_repo.utility import getCurrentDirectory, getStartRow, getCustodian
from utils.utility import fromExcelOrdinal, writeCsv
from investment_lookup.id_lookup import get_investment_Ids
import logging
logger = logging.getLogger(__name__)



def readHolding(ws, startRow):
	"""
	[Worksheet] ws, [Int] startRow => 
		[Iterator] holdings

	ws: the worksheet for holdings
	startRow: the row of the headers of the holdings
	holdings: an iterable object on holdings, where a holding is a dictionary
		containing the key-value pairs of a holding position.
	"""
	return map(partial(position, getHeaders(ws, startRow))
			  , generateHolding(ws, startRow+1))



def generateHolding(ws, startRow):
	"""
	A generator function.

	[Worksheet] ws, [Int] startRow => 
		An iterator for values from each line in the holding section.

	ws: the worksheet for holdings
	startRow: the row of the first line of holdings

	"""
	def endOfHolding(value):
		if isinstance(value, str) and value.startswith('Exposure total'):
			return True
		else:
			return False


	row = startRow
	while (row < ws.nrows and not endOfHolding(cellValue(ws, row, 0))):
		yield map(partial(cellValue, ws, row), range(ws.ncols))
		row = row + 1

	# if the function arrives here, the generator will be considered empty.



def getHeaders(ws, row):
	"""
	[Worksheet] ws, [Int] row => [List] headers

	A header cannot be an empty string.
	"""
	return list(takewhile(lambda v: v != ''
						 , map(partial(cellValue, ws, row), range(ws.ncols))))



def cellValue(ws, row, col):
	"""
	[WorkSheet] ws, [Int] row, [Int] col => [Object] value
	"""
	value = ws.cell_value(row, col)
	if isinstance(value, str):
		return value.strip()
	else:
		return value



def position(headers, cellValues):
	"""
	[List] headers, [Iterator] cellValues => [Dictionary] position
	"""
	return dict(zip(headers, cellValues))



def genevaPosition(portfolioId, position):
	"""
	[String] portfolioId, [Dictionary] position => [Dictionary] gPosition

	A Geneva position is a dictionary object that has the following list
	of keys:

	portfolio|custodian|date|geneva_investment_id|ISIN|bloomberg_figi|name
	|currency|quantity
	
	"""
	genevaPos = {}
	genevaPos['portfolio'] = portfolioId
	genevaPos['custodian'] = getCustodian()
	genevaPos['date'] = fromExcelOrdinal(position['Exposure Date']).strftime('%Y-%m-%d')
	genevaPos['name'] = position['Underlying Name']
	genevaPos['currency'] = position['Notional 1 Ccy']
	genevaPos['quantity'] = position['Notional 1']
	(genevaPos['geneva_investment_id'], genevaPos['ISIN'], genevaPos['bloomberg_figi']) = \
		get_investment_Ids(portfolioId, 'ISIN', position['ISIN'])
	
	return genevaPos



def dictToValues(keys, d):
	"""
	[List] keys, [Dictionary] d => [Iterator] values

	retrieve the list of values corresponding to the keys from the dictionary.
	"""
	return map(partial(getitem, d), keys)



def getOutputFileName(inputFile, outputDir, prefix):
	"""
	[String] inputFile, [String] outputDir, [String] prefix =>
		[String] output file name (with path)
	"""
	return join(outputDir, prefix + getDateFromFilename(inputFile) + '.csv')



def getDateFromFilename(inputFile):
	"""
	[String] inputFile => [String] date (yyyy-mm-dd)

	inputFile filename looks like (after stripping path):

	Repo Exposure Trades and Collateral Position by Agreement_CN LFAM HKH CRHKH_24_04_2019_01_00_08.xlsx
	"""
	dateString = inputFile.split('\\')[-1].split('.')[0].split('_')[0]
	return dateString[4:8] + '-' + dateString[2:4] + '-' + dateString[0:2]



def toCsv(portfolio, inputFile, outputDir, prefix):
	"""
	[String] portfolio,	[String] intputFile, [String] outputDir, [String] prefix
		=> [String] outputFile

	Side effect: create an output csv file
	"""
	gPositions = map(partial(genevaPosition, portfolio)
					, readHolding(open_workbook(inputFile).sheet_by_index(0)
								 , getStartRow()))
	headers = ['portfolio', 'custodian', 'date', 'geneva_investment_id',
				'ISIN', 'bloomberg_figi', 'name', 'currency', 'quantity']
	rows = map(partial(dictToValues, headers), gPositions)
	outputFile = getOutputFileName(inputFile, outputDir, prefix)
	writeCsv(outputFile, chain([headers], rows), '|')
	return outputFile



if __name__ == '__main__':
	import logging.config
	logging.config.fileConfig('logging.config', disable_existing_loggers=False)

		