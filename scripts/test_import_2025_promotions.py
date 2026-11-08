"""Zero monetary values do not determine a promotion's VAT rate."""
import unittest
from decimal import Decimal
from prepare_import_2025_20261109 import import_rate_candidates

class PromotionRate(unittest.TestCase):
    def test_zero_promotion_uses_separately_verified_rate(self):
        self.assertEqual(import_rate_candidates(Decimal(0),Decimal(0),True,8),[8])
    def test_zero_without_promotion_or_verified_rate_remains_ambiguous(self):
        self.assertEqual(len(import_rate_candidates(Decimal(0),Decimal(0),False,8)),4)
        self.assertEqual(len(import_rate_candidates(Decimal(0),Decimal(0),True,None)),4)
    def test_confirmation_cannot_override_nonzero_tax_amounts(self):
        self.assertEqual(import_rate_candidates(Decimal(1000),Decimal(80),True,10),[8])

if __name__=='__main__':unittest.main()
