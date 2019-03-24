import xlwings as xw
import pandas as pd
import numpy as np

from general.base import XlsRange

def generate_schedule_template(filepath, df_months):
    sheetname = 'TEST'
    wb_generic = xw.Book(filepath)
    work_sheet = wb_generic.sheets[sheetname]
    
    rm = XlsRange(1, 1, 149, 38)

    val_months = df_months.values
    work_sheet.range((rm.frow, rm.fcol), (rm.lrow, rm.lcol)).value = val_months
    wb_generic.save()


