# coding=utf-8
# 

import unittest2
from xlrd import open_workbook
from cmbhk.cmb import readHolding
from cmbhk.utility import getCurrentDirectory, getStartRow
from os.path import join



class TestCMBHK(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestCMBHK, self).__init__(*args, **kwargs)



    def testReadHolding(self):
        inputFile = join(getCurrentDirectory(), 'samples', \
                        'holding _ 16032017.xlsx')
        wb = open_workbook(inputFile)
        ws = wb.sheet_by_index(0)
        holding = list(readHolding(ws, getStartRow()))
        self.assertEqual(11, len(holding))
        self.verifyHolding1(holding[0])
        self.verifyHolding2(holding[10])



    # def testGenevaPosition(self):
    #     """
    #     This test is NOT valid in production environment. Because in production
    #     environment the investment_lookup.id_lookup.get_investment_Ids() 
    #     function uses a different HTM bond list.
    #     """
    #     inputFile = join(getCurrentDirectory(), 'samples', \
    #                         'Repo Exposure Trades and Collateral Position.xlsx')
    #     wb = open_workbook(inputFile)
    #     ws = wb.sheet_by_index(0)
    #     gPositions = list(map(partial(genevaPosition, '40002')
    #                              , readHolding(ws, getStartRow())))
    #     self.assertEqual(12, len(gPositions))
    #     self.verifyGenevaHolding1(gPositions[0])    # HTM holding
    #     self.verifyGenevaHolding2(gPositions[10])   # non HTM holding




    def verifyHolding1(self, holding):
        self.assertEqual(10, len(holding))
        self.assertAlmostEqual(49.75, holding['Indicative Price'])
        self.assertEqual('HKD', holding['Ccy'])
        self.assertEqual(3000000, holding['Settled Quantity'])
        self.assertEqual('HK0000069689', holding['Securities Identifier'])
        self.assertEqual('USD', holding['Base Ccy'])



    def verifyHolding2(self, holding):
        self.assertEqual(10, len(holding))
        self.assertAlmostEqual(100.17, holding['Indicative Price'])
        self.assertEqual('USD', holding['Ccy'])
        self.assertEqual(200000000, holding['Traded Quantity'])
        self.assertEqual('US9128282988', holding['Securities Identifier'])
        self.assertEqual('US TREASURY N/B 2.25% 02/15/2027', holding['Securities Name'])


    # def verifyHolding2(self, holding):
    #     self.assertEqual(20, len(holding))
    #     self.assertAlmostEqual('Rev Repo', holding['Product ID'])
    #     self.assertEqual('USD', holding['Notional 1 Ccy'])
    #     self.assertEqual(1500000, holding['Notional 1'])
    #     self.assertEqual('CDBLFD 0 07/18/21', holding['Underlying Name'])



    # def verifyGenevaHolding1(self, position):
    #     self.assertEqual(9, len(position))
    #     self.assertEqual('40002', position['portfolio'])
    #     self.assertEqual('HSBC-REPO', position['custodian'])
    #     self.assertEqual('2019-04-23', position['date'])
    #     self.assertEqual('USD', position['currency'])
    #     self.assertEqual(2000000, position['quantity'])
    #     self.assertEqual('XS1917106061 HTM', position['geneva_investment_id'])



    # def verifyGenevaHolding2(self, position):
    #     self.assertEqual(9, len(position))
    #     self.assertEqual('40002', position['portfolio'])
    #     self.assertEqual('ORIEAS 0 12/21/20', position['name'])
    #     self.assertEqual('2019-04-23', position['date'])
    #     self.assertEqual('XS9999998888', position['ISIN'])
    #     self.assertEqual(1000000, position['quantity'])
    #     self.assertEqual('', position['geneva_investment_id'])
