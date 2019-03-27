import os
import shutil
import xlwings as xw
import pandas as pd
import numpy as np

from src.general.base import XlsRange


def create_new_xlsm(filename):
    dir_path = os.path.dirname(os.path.realpath(__file__))
    dir_path = dir_path[:-4] # TODO: HARDCODED
    filepath = os.path.join(dir_path, '_BASE.xlsm')
    shutil.copy(filepath, filename)

def generate_schedule_template(filepath, filename, df_months):
    filepath = os.path.join(filepath, filename)
    create_new_xlsm(filepath)

    app = xw.App(visible=False)
    wb_generic = app.books.open(filepath)
    sheetname = 'MASTER'
    work_sheet = wb_generic.sheets[sheetname]
    
    rm = XlsRange(1, 1, 149, 38)

    val_months = df_months.values
    work_sheet.range((rm.frow, rm.fcol), (rm.lrow, rm.lcol)).value = val_months

    wb_generic.save()
    wb_generic.close()
    app.quit()
