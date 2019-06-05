# coding=utf-8
# 

import unittest2
from functools import partial
from xlrd import open_workbook
from cmbhk.cmb import readHolding, genevaPosition, readCash
from cmbhk.utility import getCurrentDirectory, getStartRow
from os.path import join



class TestCMBHK(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestCMBHK, self).__init__(*args, **kwargs)



    def testReadHolding(self):
        inputFile = join(getCurrentDirectory(), 'samples', \
                        'holding _ 31052019.XLS')
        wb = open_workbook(inputFile)
        ws = wb.sheet_by_index(0)
        holding = list(readHolding(ws, getStartRow()))
        self.assertEqual(2, len(holding))
        self.verifyHolding1(holding[0])
        self.verifyHolding2(holding[1])



    def testGenevaPosition(self):
        inputFile = join(getCurrentDirectory(), 'samples', \
                        'holding _ 31052019.XLS')
        wb = open_workbook(inputFile)
        ws = wb.sheet_by_index(0)
        gPositions = list(map(partial(genevaPosition, '40017', '2017-05-31') 
                             , readHolding(ws, getStartRow())))
        self.assertEqual(2, len(gPositions))
        self.verifyGenevaHolding1(gPositions[0])



    def testReadCash(self):
        inputFile = join(getCurrentDirectory(), 'samples', \
                        'cash _ 31052019.XLS')
        wb = open_workbook(inputFile)
        ws = wb.sheet_by_index(0)
        cash = dict(readCash(ws))
        self.assertEqual(len(cash), 2)
        self.assertEqual(cash['HKD'], 0)
        self.assertAlmostEqual(cash['USD'], 5338350.56)



    def verifyHolding1(self, holding):
        self.assertEqual(10, len(holding))
        self.assertAlmostEqual(103.497, holding['Indicative Price'])
        self.assertEqual('USD', holding['Ccy'])
        self.assertEqual(200000, holding['Settled Quantity (Available Balance)'])
        self.assertEqual('CH1234567890', holding['Securities Identifier'])
        self.assertEqual('USD', holding['Base Ccy'])



    def verifyHolding2(self, holding):
        self.assertEqual(10, len(holding))
        self.assertAlmostEqual(88.765, holding['Indicative Price'])
        self.assertEqual('USD', holding['Ccy'])
        self.assertEqual(200000, holding['Traded Quantity (Ledger Balance)'])
        self.assertEqual('XS1234567980', holding['Securities Identifier'])
        self.assertEqual('DFG', holding['Securities Name'])



    def verifyGenevaHolding1(self, position):
        self.assertEqual(9, len(position))
        self.assertEqual('40017', position['portfolio'])
        self.assertEqual('CMBHK', position['custodian'])
        self.assertEqual('2017-05-31', position['date'])
        self.assertEqual('USD', position['currency'])
        self.assertEqual(300000, position['quantity'])
        self.assertEqual('ABC', position['name'])
        self.assertEqual('CH1234567890', position['ISIN'])
