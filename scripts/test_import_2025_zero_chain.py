"""Regression: shared sales must not be blocked by an already-zero invoice chain."""
import unittest
from prepare_import_2025_20261109 import chain_is_verified_zero

class ZeroChainGate(unittest.TestCase):
    def test_all_three_verified_zero_needs_no_invoice_mutation(self):
        self.assertTrue(chain_is_verified_zero({'balance':[0,0.0,'0'],'issues':[]}))

    def test_unknown_amount_is_not_zero(self):
        self.assertFalse(chain_is_verified_zero({'balance':[0,None,0],'issues':[]}))
        self.assertFalse(chain_is_verified_zero({'balance':[0,'',0],'issues':[]}))
        self.assertFalse(chain_is_verified_zero({'balance':[0,' ',0],'issues':[]}))
        self.assertFalse(chain_is_verified_zero({'balance':None,'issues':[]}))

    def test_residual_or_invalid_status_still_requires_review(self):
        self.assertFalse(chain_is_verified_zero({'balance':[0,1,1],'issues':[]}))
        self.assertFalse(chain_is_verified_zero({'balance':[0,0,0],'issues':['CQT từ chối']}))

if __name__=='__main__':unittest.main()
