import numpy as np
from xlwings import Workbook, Range

def rand_numbers():
    """ produces standard normally distributed random numbers with shape (n,n)"""
    wb = Workbook.caller()  # Creates a reference to the calling Excel file
    n = int(Range('Sheet1', 'B1').value)  # Write desired dimensions into Cell B1
    rand_num = np.random.randn(n, n)
    Range('Sheet1', 'C3').value = rand_num

