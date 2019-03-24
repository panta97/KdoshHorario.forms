import calendar

from general.base import XlsRange


def get_days_ranges(start_at_zero=False):
    z = 0 if start_at_zero == False else 1
    arr = []

    # RANGE FIRST DAY BOX
    frow = 2 - z
    lrow = 25 - z
    fcol = 5 - z
    lcol = 8 - z

    for i_y in range(1, 7):
        aux_arr = []
        for i_x in range(1, 8):
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


def get_array_month(year, month):
    week_arr = calendar.monthcalendar(year, month)
    if len(week_arr) < 6:
        off_set = [0,0,0,0,0,0,0]
        week_arr.append(off_set)
    return week_arr


def transform_df(year, month, df_months):
    # GET THE RANGE POSITIONS OF ALL DAYS IN EXCEL
    range_months = get_days_ranges(start_at_zero=True)

    # GET CALENDAR ARRAY OF MONTHS
    arr_months = get_array_month(year, month)

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

