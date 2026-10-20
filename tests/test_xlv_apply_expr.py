import unittest
from core import UnitTestCreator
from core import TestSerializer
from core import debugger
from excel_enhancer import xlv_apply_expr

debugger.testing = True


class MyTestForXlvApplyExpr(unittest.TestCase):

    def test_xlv_apply_expr0(self):
        # get arguments
        arg0 = [['job', 'company', 'ssn'], ['Patent attorney', 'Anderson, Rivera and Shaffer', '521-15-1313'],
                ['Engineer, automotive', 'Font Inc', '558-10-8908'], ['Domoticien', 'Charles S.A.R.L.', '601-32-6521'],
                ['Technicien des industries du verre', 'Petitjean', '601-97-4157']]
        arg1 = 'x = x.upper() if irow==0 else x'
        arg2 = None
        arg3 = False

        # get expected result
        ret_cmp = ([['JOB', 'COMPANY', 'SSN'], ['Patent attorney', 'Anderson, Rivera and Shaffer', '521-15-1313'],
                    ['Engineer, automotive', 'Font Inc', '558-10-8908'],
                    ['Domoticien', 'Charles S.A.R.L.', '601-32-6521'],
                    ['Technicien des industries du verre', 'Petitjean', '601-97-4157']], [])
        # evaluate function
        ret = xlv_apply_expr(arg0, arg1, arg2, arg3)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))

    def test_xlv_apply_expr1(self):
        # get arguments
        arg0 = [['job', 'company', 'ssn'], ['Patent attorney', 'Anderson, Rivera and Shaffer', '521-15-1313'],
                ['Engineer, automotive', 'Font Inc', '558-10-8908'], ['Domoticien', 'Charles S.A.R.L.', '601-32-6521'],
                ['Technicien des industries du verre', 'Petitjean', '601-97-4157']]
        arg1 = "x=f'({irow},{icol}<{column}>) = {x}'"
        arg2 = None
        arg3 = False

        # get expected result
        ret_cmp = ([['(0,0<0>) = job', '(0,1<1>) = company', '(0,2<2>) = ssn'],
                    ['(1,0<0>) = Patent attorney', '(1,1<1>) = Anderson, Rivera and Shaffer', '(1,2<2>) = 521-15-1313'],
                    ['(2,0<0>) = Engineer, automotive', '(2,1<1>) = Font Inc', '(2,2<2>) = 558-10-8908'],
                    ['(3,0<0>) = Domoticien', '(3,1<1>) = Charles S.A.R.L.', '(3,2<2>) = 601-32-6521'],
                    ['(4,0<0>) = Technicien des industries du verre', '(4,1<1>) = Petitjean',
                     '(4,2<2>) = 601-97-4157']], [])
        # evaluate function
        ret = xlv_apply_expr(arg0, arg1, arg2, arg3)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))

    def test_xlv_apply_expr2(self):
        # get arguments
        arg0 = [['job', 'company', 'ssn'], ['Patent attorney', 'Anderson, Rivera and Shaffer', '521-15-1313'],
                ['Engineer, automotive', 'Font Inc', '558-10-8908'], ['Domoticien', 'Charles S.A.R.L.', '601-32-6521'],
                ['Technicien des industries du verre', 'Petitjean', '601-97-4157']]
        arg1 = "x=f'({irow},{icol}<{column}>) = {x}, args: {args}'"
        arg2 = ['arg1', 'arg2']
        arg3 = True

        # get expected result
        ret_cmp = (TestSerializer.deserialize((['job', 'company', 'ssn'], [
            ["(0,0<job>) = Patent attorney, args: ['arg1', 'arg2']",
             "(0,1<company>) = Anderson, Rivera and Shaffer, args: ['arg1', 'arg2']",
             "(0,2<ssn>) = 521-15-1313, args: ['arg1', 'arg2']"],
            ["(1,0<job>) = Engineer, automotive, args: ['arg1', 'arg2']",
             "(1,1<company>) = Font Inc, args: ['arg1', 'arg2']", "(1,2<ssn>) = 558-10-8908, args: ['arg1', 'arg2']"],
            ["(2,0<job>) = Domoticien, args: ['arg1', 'arg2']",
             "(2,1<company>) = Charles S.A.R.L., args: ['arg1', 'arg2']",
             "(2,2<ssn>) = 601-32-6521, args: ['arg1', 'arg2']"],
            ["(3,0<job>) = Technicien des industries du verre, args: ['arg1', 'arg2']",
             "(3,1<company>) = Petitjean, args: ['arg1', 'arg2']",
             "(3,2<ssn>) = 601-97-4157, args: ['arg1', 'arg2']"]]), 'DataFrame'), [])
        # evaluate function
        ret = xlv_apply_expr(arg0, arg1, arg2, arg3)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))

    def test_xlv_apply_expr3(self):
        # get arguments
        arg0 = [['job', 'company', 'ssn'], ['Patent attorney', 'Anderson, Rivera and Shaffer', '521-15-1313'],
                ['Engineer, automotive', 'Font Inc', '558-10-8908'], ['Domoticien', 'Charles S.A.R.L.', '601-32-6521'],
                ['Technicien des industries du verre', 'Petitjean', '601-97-4157']]
        arg1 = "temp = str(x)\nx = 'UPPER: ' + temp.upper()\nx += '; LOWER: ' + temp.lower()\nx += '; VARIABLE: ' + args[0]\n"
        arg2 = 'mi variable'
        arg3 = True

        # get expected result
        ret_cmp = (TestSerializer.deserialize((['job', 'company', 'ssn'], [
            ['UPPER: PATENT ATTORNEY; LOWER: patent attorney; VARIABLE: mi variable',
             'UPPER: ANDERSON, RIVERA AND SHAFFER; LOWER: anderson, rivera and shaffer; VARIABLE: mi variable',
             'UPPER: 521-15-1313; LOWER: 521-15-1313; VARIABLE: mi variable'],
            ['UPPER: ENGINEER, AUTOMOTIVE; LOWER: engineer, automotive; VARIABLE: mi variable',
             'UPPER: FONT INC; LOWER: font inc; VARIABLE: mi variable',
             'UPPER: 558-10-8908; LOWER: 558-10-8908; VARIABLE: mi variable'],
            ['UPPER: DOMOTICIEN; LOWER: domoticien; VARIABLE: mi variable',
             'UPPER: CHARLES S.A.R.L.; LOWER: charles s.a.r.l.; VARIABLE: mi variable',
             'UPPER: 601-32-6521; LOWER: 601-32-6521; VARIABLE: mi variable'], [
                'UPPER: TECHNICIEN DES INDUSTRIES DU VERRE; LOWER: technicien des industries du verre; VARIABLE: mi variable',
                'UPPER: PETITJEAN; LOWER: petitjean; VARIABLE: mi variable',
                'UPPER: 601-97-4157; LOWER: 601-97-4157; VARIABLE: mi variable']]), 'DataFrame'), [])
        # evaluate function
        ret = xlv_apply_expr(arg0, arg1, arg2, arg3)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))


if __name__ == '__main__':
    unittest.main()
