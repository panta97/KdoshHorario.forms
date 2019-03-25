import xlwings as xw
import pandas as pd
import numpy as np

from general.base import XlsRange


def read_book(filepath):
    sheetname = 'PLANTILLA'
    wb_generic = xw.Book(filepath)
    work_sheet = wb_generic.sheets[sheetname]

    # RANGE ALL MONTHS (EXTRACT ALL)
    rm = XlsRange(1, 1, 149, 38)

    # RANGE ALL INDICES
    # ri = XlsRange(1, 1, 149, 4)

    # EXTRACT RM AS FORMULAS
    arr_fml = np.array(work_sheet.range((rm.frow, rm.fcol), (rm.lrow, rm.lcol)).formula)
    df_months = pd.DataFrame(arr_fml)

    # EXTRACT RI AS VALUES
    # df_indices = work_sheet.range((ri.frow, ri.fcol), (ri.lrow, ri.lcol)).options(pd.DataFrame, empty=np.nan).value

    return df_months