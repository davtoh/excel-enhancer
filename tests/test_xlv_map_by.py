import unittest
from core import UnitTestCreator
from core import TestSerializer
from core import debugger
from excel_enhancer import xlv_map_by

debugger.testing = True


class MyTestForXlvMapBy(unittest.TestCase):

    def test_xlv_map_by0(self):  # FIXME why this automatic test fails
        # get arguments
        arg0 = [['DATOS'], [None], ['original 1'], ['original 2'], ['valor 1'], ['valor 1'], ['valor 2'], [None],
                ['valor 3']]
        arg1 = [['MAPA ENTRADA'], ['original'], ['cambio'], ['valor 1'], ['VACIO'], ['original 1'], [None], ['valor 2'],
                ['valor 3']]
        arg2 = [['MAPA SALIDA'], ['cambio'], ['cambio2'], ['valor 1 x 2'], ['ERA VACIO'], ['era original 1'], ['Vacío'],
                [None], ['valor 3']]
        arg3 = True
        arg4 = True
        arg5 = 'ignore'

        # get expected result
        ret_cmp = (TestSerializer.deserialize((['DATOS'], [[None], ['era original 1'], ['original 2'], ['valor 1 x 2'],
                                                           ['valor 1 x 2'], [None], [None], ['valor 3']]), 'DataFrame'),
                   [TestSerializer.deserialize((['DATOS', 'DATOS'], [[None, None], ['original 1', 'era original 1'],
                                                                     ['original 2', 'original 2'],
                                                                     ['valor 1', 'valor 1 x 2'],
                                                                     ['valor 1', 'valor 1 x 2'], ['valor 2', None],
                                                                     [None, None], ['valor 3', 'valor 3']]),
                                               'DataFrame')])
        # evaluate function
        ret = xlv_map_by(arg0, arg1, arg2, arg3, arg4, arg5)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))

    def test_xlv_map_by1(self):
        # get arguments
        arg0 = [['DATOS'], [None], ['original 1'], ['original 2'], ['valor 1'], ['valor 1'], ['valor 2']]
        arg1 = [['MAPA ENTRADA'], ['original'], ['cambio'], ['valor 1'], [None], ['original 1']]
        arg2 = [['MAPA SALIDA'], ['cambio'], ['cambio2'], ['valor 1 x 2'], ['vacío'], ['era original 1']]
        arg3 = True
        arg4 = True
        arg5 = None

        # get expected result
        ret_cmp = (TestSerializer.deserialize(
            (['DATOS'], [['vacío'], ['era original 1'], ['original 2'], ['valor 1 x 2'], ['valor 1 x 2'], ['valor 2']]),
            'DataFrame'), [TestSerializer.deserialize((['DATOS', 'DATOS'],
                                                       [[None, 'vacío'], ['original 1', 'era original 1'],
                                                        ['original 2', 'original 2'], ['valor 1', 'valor 1 x 2'],
                                                        ['valor 1', 'valor 1 x 2'], ['valor 2', 'valor 2']]),
                                                      'DataFrame')])
        # evaluate function
        ret = xlv_map_by(arg0, arg1, arg2, arg3, arg4, arg5)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))

    def test_xlv_map_by2(self):
        # get arguments
        arg0 = [['DATOS'], [None], ['original 1'], ['original 2'], ['valor 1'], ['valor 1'], ['valor 2'], [None],
                ['valor 3']]
        arg1 = [['MAPA ENTRADA'], ['original'], ['cambio'], ['valor 1'], ['VACIO'], ['original 1'], [None], ['valor 2'],
                ['valor 3']]
        arg2 = [['MAPA SALIDA'], ['cambio'], ['cambio2'], ['valor 1 x 2'], ['ERA VACIO'], ['era original 1'], ['Vacío'],
                [None], ['valor 3']]
        arg3 = True
        arg4 = None
        arg5 = 'ignore'

        # get expected result
        ret_cmp = (TestSerializer.deserialize((['DATOS'],
                                               [[None], ['era original 1'], ['#N/A'], ['valor 1 x 2'], ['valor 1 x 2'],
                                                [None], [None], ['valor 3']]), 'DataFrame'), [
                       TestSerializer.deserialize((['DATOS', 'DATOS'], [[None, None], ['original 1', 'era original 1'],
                                                                        ['original 2', '#N/A'],
                                                                        ['valor 1', 'valor 1 x 2'],
                                                                        ['valor 1', 'valor 1 x 2'], ['valor 2', None],
                                                                        [None, None], ['valor 3', 'valor 3']]),
                                                  'DataFrame')])
        # evaluate function
        ret = xlv_map_by(arg0, arg1, arg2, arg3, arg4, arg5)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))

    def test_xlv_map_by3(self):
        # get arguments
        arg0 = [['DATOS'], [None], ['original 1'], ['original 2'], ['valor 1'], ['valor 1'], ['valor 2'],
                ['MAPA ENTRADA']]
        arg1 = [['MAPA ENTRADA'], ['original'], ['cambio']]
        arg2 = [['MAPA SALIDA'], ['cambio'], ['cambio2']]
        arg3 = True
        arg4 = None
        arg5 = None

        # get expected result
        ret_cmp = (
        TestSerializer.deserialize((['DATOS'], [['#N/A'], ['#N/A'], ['#N/A'], ['#N/A'], ['#N/A'], ['#N/A'], ['#N/A']]),
                                   'DataFrame'), [TestSerializer.deserialize((['DATOS', 'DATOS'],
                                                                              [[None, '#N/A'], ['original 1', '#N/A'],
                                                                               ['original 2', '#N/A'],
                                                                               ['valor 1', '#N/A'], ['valor 1', '#N/A'],
                                                                               ['valor 2', '#N/A'],
                                                                               ['MAPA ENTRADA', '#N/A']]),
                                                                             'DataFrame')])
        # evaluate function
        ret = xlv_map_by(arg0, arg1, arg2, arg3, arg4, arg5)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))

    def test_xlv_map_by4(self):
        # get arguments
        arg0 = [['DATOS'], [None], ['original 1'], ['original 2'], ['valor 1'], ['valor 1'], ['valor 2']]
        arg1 = [['MAPA ENTRADA'], ['original'], ['cambio'], ['valor 1'], [None], ['original 1']]
        arg2 = [['MAPA SALIDA'], ['cambio'], ['cambio2'], ['valor 1 x 2'], ['vacío'], ['era original 1']]
        arg3 = True
        arg4 = None
        arg5 = None

        # get expected result
        ret_cmp = (TestSerializer.deserialize(
            (['DATOS'], [['vacío'], ['era original 1'], ['#N/A'], ['valor 1 x 2'], ['valor 1 x 2'], ['#N/A']]),
            'DataFrame'), [TestSerializer.deserialize((['DATOS', 'DATOS'],
                                                       [[None, 'vacío'], ['original 1', 'era original 1'],
                                                        ['original 2', '#N/A'], ['valor 1', 'valor 1 x 2'],
                                                        ['valor 1', 'valor 1 x 2'], ['valor 2', '#N/A']]),
                                                      'DataFrame')])
        # evaluate function
        ret = xlv_map_by(arg0, arg1, arg2, arg3, arg4, arg5)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))

    def test_xlv_map_by5(self):
        # get arguments
        arg0 = [['DATOS'], [None], ['original 1'], ['original 2'], ['valor 1'], ['valor 1'], ['valor 2']]
        arg1 = [['MAPA ENTRADA'], ['original'], ['cambio']]
        arg2 = [['MAPA SALIDA'], ['cambio'], ['cambio2']]
        arg3 = True
        arg4 = True
        arg5 = None

        # get expected result
        ret_cmp = (TestSerializer.deserialize(
            (['DATOS'], [[None], ['original 1'], ['original 2'], ['valor 1'], ['valor 1'], ['valor 2']]), 'DataFrame'),
                   [TestSerializer.deserialize((['DATOS', 'DATOS'], [[None, None], ['original 1', 'original 1'],
                                                                     ['original 2', 'original 2'],
                                                                     ['valor 1', 'valor 1'], ['valor 1', 'valor 1'],
                                                                     ['valor 2', 'valor 2']]), 'DataFrame')])
        # evaluate function
        ret = xlv_map_by(arg0, arg1, arg2, arg3, arg4, arg5)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))

    def test_xlv_map_by6(self):
        # get arguments
        arg0 = [['DATOS'], [None], ['original 1'], ['original 2'], ['valor 1'], ['valor 1'], ['valor 2'],
                ['MAPA ENTRADA']]
        arg1 = [['MAPA ENTRADA'], ['original'], ['cambio']]
        arg2 = [['MAPA SALIDA'], ['cambio'], ['cambio2']]
        arg3 = False
        arg4 = None
        arg5 = None

        # get expected result
        ret_cmp = ([['#N/A'], ['#N/A'], ['#N/A'], ['#N/A'], ['#N/A'], ['#N/A'], ['#N/A'], ['MAPA SALIDA']], [
            TestSerializer.deserialize(([0, 0], [['DATOS', '#N/A'], [None, '#N/A'], ['original 1', '#N/A'],
                                                 ['original 2', '#N/A'], ['valor 1', '#N/A'], ['valor 1', '#N/A'],
                                                 ['valor 2', '#N/A'], ['MAPA ENTRADA', 'MAPA SALIDA']]), 'DataFrame')])
        # evaluate function
        ret = xlv_map_by(arg0, arg1, arg2, arg3, arg4, arg5)
        # compare expected an result
        self.assertEqual(TestSerializer.serialize(ret_cmp), TestSerializer.serialize(ret))


if __name__ == '__main__':
    unittest.main()
