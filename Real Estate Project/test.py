import unittest
from unittest.mock import patch
import calculator


# Test input file
class TestRealEstateFlippingCalculator(unittest.TestCase):
    def test_input_file(self):
        test_input_file = 'test_input.txt'
        with open(test_input_file, 'r') as file:
            test_inputs = file.read().splitlines()

        with patch('builtins.input', side_effect=test_inputs):
            calculator.profit_calculations()

    if __name__ == '__main__':
        unittest.main()

