import unittest
import numpy as np
from extensions.udf_geometry import standarizePoints, polygonArea_calcule, PolyArea
from extensions.udf_json_support import xf_json_reader, xf_json_writer


class TestExtensions(unittest.TestCase):

    def test_standarize_points(self):
        pts = [(0, 0), (0, 4), (3, 0)]
        std = standarizePoints(pts)
        self.assertIsInstance(std, np.ndarray)
        self.assertEqual(std.shape, (3, 2))

    def test_polygon_area_triangle(self):
        # Right triangle with base 3, height 4 -> area = 6
        pts = [(0, 0), (0, 4), (3, 0)]
        area = polygonArea_calcule(pts)
        self.assertAlmostEqual(area, 6.0)

    def test_polygon_area_square(self):
        # Square with side 2 -> area = 4
        pts = [(0, 0), (2, 0), (2, 2), (0, 2)]
        area = polygonArea_calcule(pts)
        self.assertAlmostEqual(area, 4.0)

    def test_poly_area(self):
        x = np.array([0, 2, 2, 0])
        y = np.array([0, 0, 2, 2])
        self.assertAlmostEqual(PolyArea(x, y), 4.0)

    def test_json_reader_writer(self):
        table = [['A', 'B'], [1, 2]]
        json_str = xf_json_writer(table)
        self.assertIsInstance(json_str, str)
        recovered = xf_json_reader(json_str)
        self.assertEqual(recovered, table)


if __name__ == '__main__':
    unittest.main()
