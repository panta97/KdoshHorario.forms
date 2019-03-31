from src.etl.extract import read_book
from src.etl.transform import transform_df
from src.etl.load import generate_schedule_templates


def generate_template(path, year, month):

    # EXTRACT
    e_df_months = read_book()

    # TRANSFORM
    areas_id = [1, 2, 3]
    arr_t_df_months = []
    for area_id in areas_id:
        t_df_months = transform_df(year, month, area_id, e_df_months)
        aux_df = t_df_months.copy()
        arr_t_df_months.append(aux_df)
    
    # LOAD
    filename = '{year}-{month}.xlsm'.format(year=year, month=month)
    filepath_new = path

    # TODO: AREAS HARDCODED TEMP
    areas = [
        [1, 'DAMAS'],
        [2, 'CABALLEROS'],
        [3, 'HOMEKIDS']
    ]

    generate_schedule_templates(filepath_new, filename, arr_t_df_months, areas)

if __name__ == "__main__":
    path = r'C:\Users\roosevelt\Desktop'
    year = 2019
    month = 4
    generate_template(path, year, month)
