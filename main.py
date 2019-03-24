from etl.extract import read_book
from etl.transform import transform_df
from etl.load import generate_schedule_template


if __name__ == "__main__":
    filepath = r'/Users/roosevelt/Desktop/Horario-Maestro.xlsm'
    e_df_months = read_book(filepath)
    t_df_months = transform_df(2018, 12, e_df_months)
    generate_schedule_template(filepath, t_df_months)
