"""
if functions are placed in here while running then no data will be added until restarted
"""

import xlwings as xw
import json


@xw.func
@xw.arg('json_data', doc='json data to convert to table')
@xw.ret(expand='table', header=False, index=False, empty='NA')  #
def xf_json_reader(json_data):
    data = json.loads(json_data)
    if isinstance(data, dict):
        return
    elif isinstance(data, list):
        return data
    else:
        return data


@xw.func
@xw.arg('table', ndim=2, doc='array to json')
def xf_json_writer(table):
    data = json.dumps(table)
    return data
