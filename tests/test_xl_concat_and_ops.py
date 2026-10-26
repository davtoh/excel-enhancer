import unittest
import pandas as pd
from core import debugger
from excel_enhancer import (
    xl_concat,
    xv_concat_h,
    xv_concat_v,
    xl_rename,
    xl_select_by,
    xf_permutations,
)

debugger.testing = True


class TestXlConcatAndOps(unittest.TestCase):

    def test_xl_concat_horizontal(self):
        df1 = pd.DataFrame({'A': [1, 2]})
        df2 = pd.DataFrame({'B': [3, 4]})
        res, db = xl_concat(df1, df2, axis=1)
        self.assertEqual(list(res.columns), ['A', 'B'])
        self.assertEqual(len(res), 2)

    def test_xl_concat_vertical(self):
        df1 = pd.DataFrame({'A': [1, 2]})
        df2 = pd.DataFrame({'A': [3, 4]})
        res, db = xl_concat(df1, df2, axis=0)
        self.assertEqual(len(res), 4)

    def test_xv_concat_h(self):
        t1 = [[1], [2]]
        t2 = [[3], [4]]
        res, db = xv_concat_h(t1, t2)
        self.assertIsInstance(res, pd.DataFrame)
        self.assertEqual(res.shape, (2, 2))

    def test_xv_concat_v(self):
        t1 = [[1, 2]]
        t2 = [[3, 4]]
        res, db = xv_concat_v(t1, t2)
        self.assertIsInstance(res, pd.DataFrame)
        self.assertEqual(res.shape, (2, 2))

    def test_xl_rename(self):
        df = pd.DataFrame({'old_col': [10, 20], 'other': [30, 40]})
        res, db = xl_rename(df, 'old_col', 'new_col')
        self.assertIn('new_col', res.columns)
        self.assertNotIn('old_col', res.columns)
        self.assertIn('other', res.columns)

    def test_xl_select_by_column(self):
        df = pd.DataFrame({'A': [1], 'B': [2], 'C': [3]})
        res, db = xl_select_by(df, ['A', 'C'], op='column')
        self.assertEqual(list(res.columns), ['A', 'C'])

    def test_xf_permutations(self):
        columns = [['Red', 'S'], ['Blue', 'M'], ['Green', 'L']]
        # Permute orderly
        res, db = xf_permutations(columns, has_headers=False, ignore_na=False)
        self.assertEqual(len(res), 9)  # 3 * 3


if __name__ == '__main__':
    unittest.main()
