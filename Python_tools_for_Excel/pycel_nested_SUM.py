# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

"""
Simple example file showing how a spreadsheet can be translated to python
and executed
"""
import logging
import os
import sys
from datetime import datetime

from pycel import ExcelCompiler


def pycel_logging_to_console(enable=True):
    if enable:
        logger = logging.getLogger('pycel')
        logger.setLevel('INFO')

        console = logging.StreamHandler(sys.stdout)
        console.setLevel(logging.INFO)
        logger.addHandler(console)


if __name__ == '__main__':
    beginning = datetime.now()
    pycel_logging_to_console(enable=False)

    path = os.path.dirname(__file__)
    fname = os.path.join(path, "Nested_sum.xlsx")

    print("Loading %s..." % fname)

    # load & compile the file to a graph
    excel = ExcelCompiler(filename=fname)
    print("ExcelCompiler made", datetime.now() - beginning)

    columns = ['D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z']
    # columns = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U']
    # columns = ['V', 'W', 'X', 'Y', 'Z']
    # columns = ['H']

    max_rows = 3
    # max_rows = 3
    addresses = ['Sheet1!{}{}'.format(column, row) for row in range(2, max_rows) for column in columns]

    print("addresses made", datetime.now() - beginning)
    second_beginning = datetime.now()
    for address in addresses:
        evaluated_value = excel.evaluate(address)
        excel.set_value(address, evaluated_value)
        print("EVALUATED VALUE", address, evaluated_value)

    print("Evaluation done", datetime.now() - second_beginning)
    print("all done", datetime.now() - beginning)

    # # show the graph using matplotlib if installed
    # print("Plotting using matplotlib...")
    # try:
    #     excel.plot_graph()
    # except ImportError:
    #     pass
