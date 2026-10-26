import unittest
import pandas as pd
from core import debugger
from excel_enhancer import xl_format_by

debugger.testing = True


class MyTestForXlFormatBy(unittest.TestCase):

    def test_xl_format_by_without_headers(self):
        data = [['apple', 'banana'], ['carrot', 'date']]
        template = 'fruit: {}'
        ret, steps = xl_format_by(data, template, has_headers=False)
        expected = [['fruit: apple', 'fruit: banana'], ['fruit: carrot', 'fruit: date']]
        self.assertEqual(ret, expected)
        self.assertEqual(len(steps), 1)

    def test_xl_format_by_with_headers(self):
        data = pd.DataFrame([['apple', 10], ['banana', 20]], columns=['Item', 'Count'])
        template = 'Value: {}'
        ret, steps = xl_format_by(data, template, has_headers=True)
        self.assertTrue(isinstance(ret, pd.DataFrame))
        self.assertEqual(ret['Item'].tolist(), ['Value: apple', 'Value: banana'])
        self.assertEqual(ret['Count'].tolist(), ['Value: 10', 'Value: 20'])
        self.assertEqual(len(steps), 1)

    def test_xl_format_by_invalid_template(self):
        data = [['apple']]
        with self.assertRaises(ValueError):
            xl_format_by(data, 123)


if __name__ == '__main__':
    unittest.main()
