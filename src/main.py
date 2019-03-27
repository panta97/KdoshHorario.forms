from src.etl.extract import read_book
from src.etl.transform import transform_df
from src.etl.load import generate_schedule_template


def generate_template(path, year, month):
    e_df_months = read_book()
    t_df_months = transform_df(year, month, e_df_months)
    filename = '{year}-{month}.xlsm'.format(year=year, month=month)
    filepath_new = path
    generate_schedule_template(filepath_new, filename, t_df_months)

if __name__ == "__main__":
    e_df_months = read_book()
    t_df_months = transform_df(2019, 4, e_df_months)
    filename = 'NUEVO.xlsm'
    filepath_new = r'C:\Users\roosevelt\Desktop'
    generate_schedule_template(filepath_new, filename, t_df_months)
