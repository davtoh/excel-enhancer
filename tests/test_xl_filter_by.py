import unittest
import pandas as pd
from core import debugger
from excel_enhancer import xl_filter_by

debugger.testing = True


class TestXlFilterBy(unittest.TestCase):

    def setUp(self):
        self.df = pd.DataFrame({
            'Name': ['José', 'Maria', 'John', 'Carlos', 'Ana'],
            'Role': ['Manager', 'Developer', 'Designer', 'Developer', 'Lead'],
            'Score': [85, 92, 78, 95, 88]
        })

    def test_filter_exact_match(self):
        res, db = xl_filter_by(self.df, 'Role', 'Developer')
        self.assertEqual(len(res), 2)
        self.assertListEqual(res['Name'].tolist(), ['Maria', 'Carlos'])

    def test_filter_startswith(self):
        # mode 1 is startswith
        res, db = xl_filter_by(self.df, 'Role', 'De', mode=1)
        self.assertEqual(len(res), 3)  # Developer, Designer, Developer
        self.assertIn('John', res['Name'].tolist())

    def test_filter_contains(self):
        # mode 3 is contains
        res, db = xl_filter_by(self.df, 'Role', 'sign', mode=3)
        self.assertEqual(len(res), 1)
        self.assertEqual(res.iloc[0]['Name'], 'John')

    def test_filter_ignore_accents(self):
        # 'Jose' matches 'José' with ignore_accents=True
        res, db = xl_filter_by(self.df, 'Name', 'Jose', ignore_accents=True)
        self.assertEqual(len(res), 1)
        self.assertEqual(res.iloc[0]['Name'], 'José')

    def test_filter_ignore_case(self):
        # 'developer' matches 'Developer' with ignore_case=True
        res, db = xl_filter_by(self.df, 'Role', 'developer', ignore_case=True)
        self.assertEqual(len(res), 2)

    def test_filter_with_exclude(self):
        # Exclude Developer
        res, db = xl_filter_by(self.df, 'Role', include=None, exclude='Developer')
        self.assertEqual(len(res), 3)
        self.assertNotIn('Developer', res['Role'].tolist())


if __name__ == '__main__':
    unittest.main()
