from unittest import TestCase
from main import calculate_ref_nro
import csv


class TestCalculate_ref_nro(TestCase):
    def test_ref_nos(self):
        test_ref_no_path = './test_ref_nos.csv'
        test_ref_nos = csv.DictReader(open(test_ref_no_path, 'r'))
        i = 1
        for row in test_ref_nos:
            self.assertEqual(row['Numbers'], calculate_ref_nro(14, 16, i))
            i += 1
