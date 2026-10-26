import unittest
import pandas as pd
from core import debugger
from excel_enhancer import xl_group_by

debugger.testing = True


class TestXlGroupBy(unittest.TestCase):

    def setUp(self):
        self.df = pd.DataFrame({
            'Department': ['Sales', 'Sales', 'Engineering', 'Engineering', 'HR'],
            'Salary': [50000, 60000, 80000, 90000, 45000],
            'Bonus': [5000, 6000, 8000, 7000, 3000]
        })

    def test_group_by_single_column_sum_mode_1(self):
        res, db = xl_group_by(self.df, 'Department', 'Salary', 'sum', mode='1')
        self.assertIsInstance(res, pd.DataFrame)
        self.assertIn('Department', res.columns)
        self.assertIn('SUM SALARY', res.columns)
        # Check engineering sum = 170000, HR = 45000, Sales = 110000
        dept_salaries = dict(zip(res['Department'], res['SUM SALARY']))
        self.assertEqual(dept_salaries['Engineering'], 170000)
        self.assertEqual(dept_salaries['Sales'], 110000)
        self.assertEqual(dept_salaries['HR'], 45000)

    def test_group_by_mode_2(self):
        res, db = xl_group_by(self.df, 'Department', 'Salary', 'mean', mode='2')
        self.assertIsInstance(res, pd.DataFrame)
        self.assertIn('MEAN SALARY', res.columns)
        self.assertEqual(res.loc['Engineering', 'MEAN SALARY'], 85000)

    def test_group_by_multiple_selections(self):
        res, db = xl_group_by(self.df, 'Department', ['Salary', 'Bonus'], ['mean', 'sum'], mode='1')
        self.assertIsInstance(res, pd.DataFrame)
        self.assertIn('MEAN SALARY', res.columns)
        self.assertIn('SUM BONUS', res.columns)

    def test_group_by_custom_resume(self):
        df_text = pd.DataFrame({
            'Category': ['Fruit', 'Fruit', 'Veg'],
            'Item': ['Apple', 'Banana', 'Carrot']
        })
        res, db = xl_group_by(df_text, 'Category', 'Item', 'resume', mode='1', params=[", "])
        cat_items = dict(zip(res['Category'], res['RESUME ITEM']))
        self.assertEqual(cat_items['Fruit'], 'Apple, Banana')
        self.assertEqual(cat_items['Veg'], 'Carrot')


if __name__ == '__main__':
    unittest.main()
