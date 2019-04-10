import requests
from src.general.constants import PROD_BASE_URL as base_url

def get_areas_by_sede(sede_id):
    url = base_url + 'personal/excel/sede'
    querystring = {'idSede' : sede_id}
    payload = ''
    headers = {
        'cache-control': 'no-cache'
    }

    response = requests.request('GET', url, data=payload, headers=headers, params=querystring)
    json_data = response.json()
    return json_data


class Area:
    def __init__(self, ):
        self.arr_areas = None

    def call_service(self, sede_id):
        self.arr_areas = get_areas_by_sede(sede_id)

    def get_areas_id(self):
        arr_id_areas = []
        for item in self.arr_areas:
            arr_id_areas.append(item['idArea'])
        return arr_id_areas
    
    def get_areas(self):
        arr_areas = []
        for item in self.arr_areas:
            aux_arr = [item['idArea'], item['nombreArea']]
            arr_areas.append(aux_arr)
        return arr_areas

    def get_sheetnames(self):
        arr_sheetnames = []
        for item in self.arr_areas:
            arr_sheetnames.append(item['nombreArea'])
        return arr_sheetnames

    def get_namesede(self):
        return self.arr_areas[0]['nombreSede']
    