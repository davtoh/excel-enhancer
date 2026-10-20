import pandas as pd
import xlwings as xw
import numpy as np
from scipy.stats import moment
from core import helper_correct_to_str


def standarizePoints(pts, aslist=False):
    """
    converts points to a standard form
    :param pts: list or array of points
    :param aslist: True to return list instead of array
    :return: standard points

    example::

        standarizePoints((1,2))
        >>> array([[1, 2]])
        standarizePoints([[[[(1,2)]]]])
        >>> array([[1, 2]])
        standarizePoints([[[[(1,2),(1,2),(2,3)]]]],True)
        >>> [(1, 2), (1, 2), (2, 3)]
        standarizePoints([[[[(1,2,3),(1,2,3)]]]],True)
        >>> [(1, 2), (1, 2), (2, 3)]
        standarizePoints([(1,2,3,4,5),(1,2,3,4,5)],True)
        >>> [(1, 1), (2, 2), (3, 3), (4, 4), (5, 5)]
    """
    # return zip(*splitPoints(pts))
    pts = np.array(pts)
    shape = pts.shape
    if len(shape) == 1 and pts.size == 2:
        pts = np.array([pts])
        shape = pts.shape

    half = pts.size // 2

    if pts.size != 0 and (2 not in shape or half not in shape):
        # if pts.size != 0 and (2 != shape[0] and 2 != shape[-1]):
        raise Exception(
            "array of shape {} does not seems to correspond to points".format(shape))

    # return empty points
    if pts.size == 0:
        if aslist:
            return []
        return pts

    # convert to shape (None,2)
    if shape.index(2) < shape.index(half):
        pts = pts.T

    pts = pts.reshape(half, 2)
    if aslist:  # as list
        return list(map(tuple, pts))
    return pts


def polygonArea_calcule(pts):
    """
    Area of points calculating polygon Area.

    :param pts: points.
    :return: area value.

    ..note::
        - If polygon is incomplete (last is not first point) it completes the array.
        - If the polygon crosses over itself the algorithm will fail.
        - Based on http://www.mathopenref.com/coordpolygonarea.html
    """
    pts = standarizePoints(pts)
    if (pts[0] == pts[-1]).all():  # test if last is the first (closed points)
        x, y = pts[:, 0], pts[:, 1]
    else:  # complete the points (last point must be the first)
        x = pts[list(range(len(pts))) + [0], 0]
        y = pts[list(range(len(pts))) + [0], 1]
    xmul = x[:] * np.roll(y, -1)  # shifted multiplication in axis x
    ymul = y[:] * np.roll(x, -1)  # shifted multiplication in axis y
    return np.abs(np.sum(xmul) - np.sum(ymul)) / 2.  # summation and absolute area


@xw.func
@xw.arg('points_x', doc='list of x points')
@xw.arg('points_y', doc='list of y points')
#@xw.ret(expand='table', header=False, index=False, empty='NA')  #
def xv_area(points_x, points_y):

    # DONE NO_CUENTA, ID, NOMBRE, CONCEPTO, CREDITO, DEBITO
    pts = [(x, y) for x, y in zip(points_x, points_y)]
    return polygonArea_calcule(pts)


def PolyArea(x, y):
    # https://stackoverflow.com/a/30408825/5288758
    return 0.5*np.abs(np.dot(x, np.roll(y, 1))-np.dot(y, np.roll(x, 1)))