import calendar
import numpy as np
from src.service.employees import  get_employees_by_area

def get_days_ranges(start_at_zero=False):
    z = 0 if start_at_zero == False else 1
    arr = []

    # RANGE FIRST DAY BOX
    frow = 2 - z
    lrow = 25 - z
    fcol = 5 - z # STARTS FROM 5 BECAUSE THE FIRST 4 COLUMNS ARE POPULATED FROM EMPLOYEES SERVICE
    lcol = 8 - z

    for _ in range(1, 7):
        aux_arr = []
        for _ in range(1, 8):
            new_day_range = [(frow,fcol), (lrow, lcol)]
            aux_arr.append(new_day_range)
            fcol += 5
            lcol += 5
        fcol = 5 - z
        lcol = 5 - z
        frow += 25
        lrow += 25
        arr.append(aux_arr)

    return arr

def get_dayname(idx):
    arr_days = ['Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado', 'Domingo']
    return arr_days[idx]

def get_monthname(idx):
    new_idx = idx - 1
    arr_months = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto',
                  'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    return arr_months[new_idx]

def get_array_month(year, month):
    week_arr = calendar.monthcalendar(year, month)
    if len(week_arr) < 6:
        off_set = [0,0,0,0,0,0,0]
        week_arr.append(off_set)
    return week_arr

def add_employees_areas(df, area_id):
    json_response = get_employees_by_area(area_id)
    arr_data = []

    # WE RETURN THE DATAFRAME HERE BECAUSE THERE ARE NO EMPLOYEES FOR 
    # CHOSEN AREA
    if len(json_response) == 0:
        return df

    for dict_item in json_response:
        area_id = dict_item.get('idArea')
        area_name = dict_item.get('nombreArea')
        employee_id = dict_item.get('idPersonal')
        employee_name = dict_item.get('nombrePersonal').split()[0] # GET ONLY THE FIRST NAME

        arr_item = [area_id, employee_id, area_name, employee_name]
        arr_data.append(arr_item)

    number_employees = len(arr_data)
    frow = 6 - 1
    lrow = frow + number_employees
    off_set = 25 - number_employees #EMPTY ROWS EMPTY EMPLOYEES

    arr_np = np.array(arr_data)
    arr_np = arr_np.transpose()
    

    # LOOP THROUGH EVERY ROW IN SHEET

    for _ in range(1, 7):
        df.iloc[frow:lrow][0] = arr_np[0] # AREA_ID
        df.iloc[frow:lrow][1] = arr_np[1] # EMPLOYEE_ID
        df.iloc[frow:lrow][2] = arr_np[2] # AREA_NAME
        df.iloc[frow:lrow][3] = arr_np[3] # EMPLOYEE_NAME

        frow = frow + number_employees + off_set
        lrow = frow + number_employees

    return df


def fill_area_employees(df, area_id):
    # SET FIRST FOUR COLUMNS (AREA_ID, EMPLOYEE_ID, AREA_NAME, EMPLOYEE_NAME)
    df = add_employees_areas(df, area_id)
    return df

def fill_month_df(year, month, df_months):
    # GET THE RANGE POSITIONS OF ALL DAYS IN EXCEL
    range_months = get_days_ranges(start_at_zero=True)

    # GET CALENDAR ARRAY OF MONTHS
    arr_months = get_array_month(year, month)

    # SET MONTH NAME AND YEAR NUMBER
    month_name = get_monthname(month)
    title_date = '{month} - {year}'.format(month=month_name, year=year)
    df_months[4][0] = title_date

    date = 1
    for xw, week in enumerate(range_months):
        for xd, day in enumerate(week):
            is_valid_date = arr_months[xw][xd] # GET THE DAY
            frow = day[0][0]
            fcol = day[0][1]

            if is_valid_date == 0:
                # DELETE FORMULAS ON INVALID DAYS
                df_months.iloc[(frow+4):(frow+24)][fcol+3] = ''
                df_months.iloc[frow+3][(fcol):(fcol+3)] = ''
                continue

            # SET WEEKDAY AND DATE
            df_months[fcol][frow] = get_dayname(xd)
            df_months[fcol][frow + 1] = date

            # SET SHIFTS
            df_months[fcol][frow + 2] = 'Mañana'
            df_months[fcol + 1][frow + 2] = 'Tarde'
            df_months[fcol + 2][frow + 2] = 'Noche'
            date += 1

    return df_months


def transform_df(year, month, area_id, df):
    # COPY BY VALUE
    w_df = df.copy()
    w_df = fill_month_df(year, month, w_df)
    w_df = fill_area_employees(w_df, area_id)
    return w_df