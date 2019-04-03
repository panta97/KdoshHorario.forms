from src.etl.extract import read_book
from src.etl.transform import transform_df
from src.etl.load import generate_schedule_templates
from src.service.areas import Area

def generate_template(path, year, month, sede):

    # INITIALIZE AREA CLASS (CALL SERVICE INTERNALLY)
    ars = Area()
    ars.call_service(sede)

    # EXTRACT
    e_df_months = read_book()

    # TRANSFORM
    areas_id = ars.get_areas_id()
    arr_t_df_months = []
    for area_id in areas_id:
        t_df_months = transform_df(year, month, area_id, e_df_months)
        arr_t_df_months.append(t_df_months)
    
    # LOAD
    namesede = ars.get_namesede()
    filename = '{sede}-{year}-{month}.xlsm'.format(sede=namesede, year=year, month=month)
    filepath_new = path

    # TODO: AREAS HARDCODED TEMP
    areas = ars.get_areas()
    sheetnames = ars.get_sheetnames()
    generate_schedule_templates(filepath_new, filename, arr_t_df_months, areas, sheetnames)

if __name__ == "__main__":
    path = r'C:\Users\roosevelt\Desktop\test'
    year = 2019
    month = 4
    sede = 1
    generate_template(path, year, month, sede)
